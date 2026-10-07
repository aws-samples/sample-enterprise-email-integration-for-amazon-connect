"""
Amazon Connect <-> Microsoft 365 inbound email Lambda.

Triggered on a schedule by Amazon EventBridge. Polls the Microsoft 365 mailbox Inbox
through the Microsoft Graph API for unread messages and, for each one, creates an
Amazon Connect email contact via StartEmailContact. It also:
  - threads conversations by recovering the Connect ContactId from the reply's
    In-Reply-To header (passed as RelatedContactId, after confirming the sender is the
    customer on that contact) and by forwarding the real RFC 5322 Message-ID; and
  - embeds small inline images as data: URIs so they render in place, while other
    attachments are uploaded to S3 for Connect to ingest.

Required environment variables:
  SECRETS_MANAGER_SECRET_NAME  Secret holding M365 credentials
                               (tenant_id, client_id, client_secret, mailbox_email).
  S3_BUCKET_NAME               Bucket for inbound attachments.
  CONNECT_INSTANCE_ID          Amazon Connect instance id.
  CONNECT_CONTACT_FLOW_ID      Inbound contact flow id.
  DESTINATION_EMAIL_ADDRESS    Connect email address that receives the contact.
Optional:
  BATCH_SIZE           Max unread emails to process per run (default 10).
  INLINE_IMAGE_MAX_MB  Max inline-image size embedded as a data: URI (default 2);
                       larger inline images are delivered as attachments instead.

Logging rule: log only ContactIds, Graph message ids, operation names, counts, HTTP
status codes and Graph error codes. Never log subjects, filenames, S3 keys, Content-IDs,
addresses, Internet Message-IDs, response bodies, draft ids, tokens or secret values.
"""
import json
import boto3
import urllib.request
import urllib.parse
import urllib.error
import os
import re
import base64
import hashlib
import functools
from botocore.config import Config
from botocore.exceptions import ClientError

# Generate presigned URLs with modern SigV4 signing against the regional, virtual-hosted-style
# endpoint. The boto3 default can emit a legacy SigV2 / global-endpoint URL that Amazon Connect's
# attachment ingestion cannot consume (the attached file lands as 0 bytes / FAILED), so we pin it.
_S3_PRESIGN_CONFIG = Config(signature_version='s3v4', s3={'addressing_style': 'virtual'})

# Environment configuration
SECRETS_NAME = os.environ['SECRETS_MANAGER_SECRET_NAME']
S3_BUCKET = os.environ.get('S3_BUCKET_NAME')
BATCH_SIZE = int(os.environ.get('BATCH_SIZE', '10'))
CONNECT_INSTANCE_ID = os.environ['CONNECT_INSTANCE_ID']
CONNECT_CONTACT_FLOW_ID = os.environ['CONNECT_CONTACT_FLOW_ID']
DESTINATION_EMAIL_ADDRESS = os.environ['DESTINATION_EMAIL_ADDRESS']
# Inline images at or below this size are embedded into the body as data: URIs so they
# render inline (matching how Connect stores agent-added images). Larger inline images
# fall back to being delivered as regular attachments to keep the message body within limits.
INLINE_IMAGE_MAX_MB = float(os.environ.get('INLINE_IMAGE_MAX_MB', '2'))

GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
HTTP_TIMEOUT_SECONDS = 20
# Only these hosts (plus *.amazonaws.com) may be contacted by urlopen.
_ALLOWED_HOSTS = frozenset({'login.microsoftonline.com', 'graph.microsoft.com'})
# Placeholder written by the CloudFormation template until the reader updates the secret.
PLACEHOLDER_VALUE = 'REPLACE_ME'


class GraphRequestError(Exception):
    """HTTP error from Graph or the token endpoint, reduced to status and Graph error code."""

    def __init__(self, status, error_code):
        self.status = status
        self.error_code = error_code
        super().__init__(f"HTTP {status} error_code={error_code}")


def _check_url(url):
    """Allow only https URLs to the Microsoft identity/Graph hosts or *.amazonaws.com."""
    parts = urllib.parse.urlsplit(url)
    host = (parts.hostname or '').lower()
    if parts.scheme != 'https' or not (host in _ALLOWED_HOSTS or host.endswith('.amazonaws.com')):
        raise ValueError('URL rejected: scheme or host not allowed')
    return url


def _graph_user_url(mailbox, *path_parts):
    """Build https://graph.microsoft.com/v1.0/users/{mailbox}/... with every segment percent-encoded."""
    segments = [urllib.parse.quote(mailbox, safe='')]
    segments.extend(urllib.parse.quote(str(part), safe='') for part in path_parts)
    return f"{GRAPH_API_BASE}/users/" + '/'.join(segments)


def _graph_error_code(http_error):
    """Extract error.code from a Graph error body without exposing the body itself."""
    try:
        body = json.loads(http_error.read().decode('utf-8'))
        return str(body.get('error', {}).get('code') or 'unknown')
    except (ValueError, AttributeError, OSError, TypeError):
        return 'unknown'


def _error_summary(exc):
    """Safe one-line description of an exception: type plus status or error code only."""
    if isinstance(exc, GraphRequestError):
        return f"GraphRequestError status={exc.status} error_code={exc.error_code}"
    if isinstance(exc, urllib.error.HTTPError):
        return f"HTTPError status={exc.code}"
    if isinstance(exc, ClientError):
        return f"ClientError code={exc.response.get('Error', {}).get('Code', 'unknown')}"
    return type(exc).__name__


@functools.lru_cache(maxsize=4)
def _related_pattern(domain):
    """Compile (and cache) the <uuid@domain> matcher for our outbound mail domain.

    Optional angle brackets, case-insensitive on both the hex UUID and the domain.
    Group 1 is the Connect ContactId.
    """
    return re.compile(
        r"<?([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})@"
        + re.escape(domain) + r">?",
        re.IGNORECASE,
    )

def http_request(url, method='GET', data=None, headers=None):
    """
    Simple HTTP request using urllib (no external dependencies).
    Only https to an allowlisted host is permitted; HTTP errors become GraphRequestError.
    """
    if headers is None:
        headers = {}
    if data and isinstance(data, dict):
        data = urllib.parse.urlencode(data).encode('utf-8')
    elif data and isinstance(data, str):
        data = data.encode('utf-8')

    _check_url(url)
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=HTTP_TIMEOUT_SECONDS) as response:  # nosec B310  # https + host allowlist enforced by _check_url
            return {'status_code': response.getcode(), 'content': response.read().decode('utf-8')}
    except urllib.error.HTTPError as e:
        raise GraphRequestError(e.code, _graph_error_code(e)) from None

def load_credentials():
    """Read the M365 credential JSON (tenant_id, client_id, client_secret, mailbox_email)."""
    secrets_client = boto3.client('secretsmanager')
    response = secrets_client.get_secret_value(SecretId=SECRETS_NAME)
    return json.loads(response['SecretString'])

def secret_not_configured(credentials):
    """True while the template placeholder is still in the secret (reader has not updated it)."""
    return any(
        (credentials.get(key) or '').strip() == PLACEHOLDER_VALUE
        for key in ('tenant_id', 'client_id', 'mailbox_email')
    )

def get_m365_token(credentials):
    """
    Retrieve an M365 access token using the OAuth2 client-credentials flow.
    Returns: (access_token, mailbox_email)
    """
    # OAuth2 token request
    tenant = urllib.parse.quote(credentials['tenant_id'], safe='')
    token_url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    token_data = {
        'grant_type': 'client_credentials',
        'client_id': credentials['client_id'],
        'client_secret': credentials['client_secret'],
        'scope': 'https://graph.microsoft.com/.default'
    }

    response = http_request(token_url, method='POST', data=token_data,
                          headers={'Content-Type': 'application/x-www-form-urlencoded'})
    return json.loads(response['content'])['access_token'], credentials['mailbox_email']

def get_unread_emails(token, mailbox_email):
    """
    Fetch unread emails from the M365 mailbox Inbox folder using the Microsoft Graph API.
    Only the Inbox is polled so Junk Email and Deleted Items never become contacts.
    """
    # Graph requires every $orderby property to also lead the $filter, in the same order
    # (otherwise InefficientFilter). The receivedDateTime predicate is always true.
    params = {
        '$filter': 'receivedDateTime ge 1970-01-01T00:00:00Z and isRead eq false',
        '$top': str(BATCH_SIZE),
        '$orderby': 'receivedDateTime desc',
        '$select': 'id,subject,from,body,receivedDateTime,internetMessageId,internetMessageHeaders'
    }

    url = _graph_user_url(mailbox_email, 'mailFolders', 'inbox', 'messages') + '?' + urllib.parse.urlencode(params)
    response = http_request(url, headers={'Authorization': f'Bearer {token}'})
    return json.loads(response['content']).get('value', [])

def process_email_attachments(token, mailbox_email, message_id):
    """
    Download email attachments from Microsoft Graph and split them into two groups:
      - Inline images (isInline with a Content-ID) small enough to embed are returned as a
        {normalized_content_id: data_uri} map, for inlining into the body so they render in
        place (the agent workspace and quoted reply history) instead of showing as broken
        cid: references.
      - Everything else (real file attachments, and oversized inline images) is uploaded to
        S3 and returned as Connect attachment objects, preserving the existing behavior.
    Returns: (attachments, inline_images)
    """
    # Get attachments from Microsoft Graph API
    url = _graph_user_url(mailbox_email, 'messages', message_id, 'attachments')
    response = http_request(url, headers={'Authorization': f'Bearer {token}'})
    
    if response['status_code'] != 200:
        return [], {}
    
    attachments = []
    inline_images = {}
    s3_client = boto3.client('s3', region_name=os.environ.get('AWS_REGION'), config=_S3_PRESIGN_CONFIG)
    
    # Process up to 10 attachments per email
    for i, attachment in enumerate(json.loads(response['content']).get('value', [])[:10]):
        if attachment.get('@odata.type') != '#microsoft.graph.fileAttachment':
            continue

        size_mb = attachment.get('size', 0) / (1024 * 1024)
        content_type = attachment.get('contentType', 'application/octet-stream')
        content_id = (attachment.get('contentId') or '').strip()
        content_bytes = attachment.get('contentBytes')

        # Inline image with a Content-ID, small enough to embed -> data: URI for the body.
        # This mirrors how Connect stores agent-added images, so it renders inline and
        # survives into outbound reply history. It is NOT added as a separate attachment.
        if (attachment.get('isInline') and content_id and content_bytes
                and content_type.lower().startswith('image/')
                and size_mb <= INLINE_IMAGE_MAX_MB):
            inline_images[_normalize_cid(content_id)] = f"data:{content_type};base64,{content_bytes}"
            print(f"Inlined image: graph_message_id={message_id} index={i} size_mb={size_mb:.2f}")
            continue

        # Everything else (files, and oversized inline images) -> S3 attachment. Skip > 20MB.
        if size_mb <= 20:
            try:
                # Decode base64 attachment content
                content = base64.b64decode(content_bytes)
                
                # Clean filename for API constraints (1-256 chars, no control chars)
                filename = attachment['name'].replace(' ', '_').replace('/', '_').replace('\\', '_')
                if len(filename) > 256:
                    filename = filename[:256]
                
                # Create S3 key with organized structure
                key = f"email-attachments/{message_id.replace('/', '_')}/{i}_{filename}"
                
                # Upload to S3 with metadata
                s3_client.put_object(
                    Bucket=S3_BUCKET,
                    Key=key,
                    Body=content,
                    ContentType=content_type,
                    Metadata={
                        'email_id': message_id.replace('/', '_'),
                        'original_name': attachment['name'],
                        'attachment_index': str(i)
                    }
                )
                
                # Generate pre-signed URL (1 hour expiry)
                presigned_url = s3_client.generate_presigned_url(
                    'get_object',
                    Params={'Bucket': S3_BUCKET, 'Key': key},
                    ExpiresIn=3600
                )
                
                # Attachment object for the StartEmailContact API
                attachment_obj = {
                    'FileName': filename,
                    'S3Url': presigned_url
                }
                attachments.append(attachment_obj)
                
                print(f"Processed attachment: graph_message_id={message_id} index={i} size_mb={size_mb:.2f}")
                
            except Exception as e:
                print(f"Failed to process attachment: graph_message_id={message_id} index={i} "
                      f"error={_error_summary(e)}")
    
    return attachments, inline_images

def extract_header(headers, name):
    """Case-insensitive lookup of a single internet message header value. Returns '' if absent."""
    target = name.lower()
    for header in headers or []:
        if header.get('name', '').lower() == target:
            return header.get('value', '') or ''
    return ''

def _normalize_cid(cid):
    """Normalize a Content-ID for matching against body cid: references (drop <>, lowercase)."""
    return (cid or '').strip().strip('<>').strip().lower()

def embed_inline_images(html, inline_images):
    """
    Replace <img src="cid:CONTENT_ID"> references in the HTML body with self-contained
    data: URIs so inbound inline images render in place (agent workspace and quoted reply
    history) instead of as broken cid: references. Unmatched cid: references are left as-is.
    """
    if not html or not inline_images:
        return html

    def _replace(match):
        prefix, raw_cid, suffix = match.group(1), match.group(2), match.group(3)
        data_uri = inline_images.get(_normalize_cid(raw_cid))
        return f"{prefix}{data_uri}{suffix}" if data_uri else match.group(0)

    return re.sub(r'(src\s*=\s*["\'])\s*cid:([^"\']+?)\s*(["\'])', _replace, html, flags=re.IGNORECASE)

def parse_related_contact_id(in_reply_to, domain):
    """
    Extract a Connect ContactId from an In-Reply-To header of the form <contactId@domain>,
    where domain is our outbound mail domain (the M365 mailbox domain). Returns the first
    (leftmost) matching UUID, or None when the header is absent/empty or has no matching token.
    """
    if not in_reply_to or not in_reply_to.strip() or not domain:
        return None
    match = _related_pattern(domain).search(in_reply_to)
    return match.group(1) if match else None

def sender_matches_contact(sender_address, contact):
    """
    Pure check used before chaining a reply to an earlier contact. Returns (matched, reason_code).
    Matches when the inbound sender address equals Contact.CustomerEndpoint.Address
    (trimmed, case-insensitive). If the earlier contact carries no customer address at all,
    the reply is still chained (the behaviour before this check existed), so threading never
    breaks for lack of data; only an explicit mismatch stops the link. Reason codes never
    contain addresses.
    """
    customer_address = ((contact or {}).get('CustomerEndpoint') or {}).get('Address') or ''
    if not customer_address.strip():
        return True, 'related_customer_address_missing'
    if (sender_address or '').strip().lower() != customer_address.strip().lower():
        return False, 'related_sender_mismatch'
    return True, 'related_sender_matched'

def verify_related_contact(connect_client, candidate_contact_id, sender_address):
    """
    Confirm a candidate RelatedContactId (recovered from In-Reply-To) belongs to the same
    customer before chaining. Anyone can forge an In-Reply-To token, so the ContactId alone
    must not link a stranger's email into an existing thread (the agent's quoted reply would
    disclose earlier history). Returns the ContactId to chain, or None to start unlinked.

    The customer address comes from DescribeContact (Contact.CustomerEndpoint.Address). If
    that field is empty, the reply is chained and reason related_customer_address_missing is
    logged; if a DescribeContact error occurs, the reply starts unlinked.
    """
    try:
        response = connect_client.describe_contact(
            InstanceId=CONNECT_INSTANCE_ID, ContactId=candidate_contact_id)
        contact = response.get('Contact', {})
    except Exception as e:
        print(f"RelatedContactId dropped: reason=related_describe_failed "
              f"candidate={candidate_contact_id} error={_error_summary(e)}")
        return None
    matched, reason = sender_matches_contact(sender_address, contact)
    if not matched:
        print(f"RelatedContactId dropped: reason={reason} candidate={candidate_contact_id}")
        return None
    if reason == 'related_customer_address_missing':
        print(f"RelatedContactId kept without sender check: reason={reason} candidate={candidate_contact_id}")
    return candidate_contact_id

def build_client_token(mailbox_email, graph_message_id):
    """
    Idempotency token for StartEmailContact: SHA-256 of mailbox and Graph message id (64 hex
    chars, API limit is 500). The same M365 message always yields the same token, so a retry
    after a failed mark-as-read does not create a second contact.
    """
    return hashlib.sha256(f"{mailbox_email}:{graph_message_id}".encode('utf-8')).hexdigest()

def build_raw_headers(email_data):
    """
    Build the StartEmailContact RAW message Headers:
    - MESSAGE_ID from the real RFC 5322 Message-ID (Graph internetMessageId), preserved exactly;
      omitted when absent/empty.
    - IN_REPLY_TO / REFERENCES forwarded from the inbound headers when non-empty.
    """
    headers = {}
    if (email_data.get('internetMessageId') or '').strip():
        headers['MESSAGE_ID'] = email_data['internetMessageId']
    internet_headers = email_data.get('internetMessageHeaders', [])
    in_reply_to = extract_header(internet_headers, 'In-Reply-To')
    if in_reply_to:
        headers['IN_REPLY_TO'] = in_reply_to
    references = extract_header(internet_headers, 'References')
    if references:
        headers['REFERENCES'] = references
    return headers

def create_connect_email_contact(email_data, attachments, related_contact_id=None, inline_images=None,
                                 client_token=None, connect_client=None):
    """
    Create email contact in Amazon Connect using StartEmailContact API
    """
    if connect_client is None:
        connect_client = boto3.client('connect')
    
    # Extract sender information
    from_email = email_data.get('from', {}).get('emailAddress', {})
    sender_email = from_email.get('address', 'unknown@domain.com')
    sender_name = from_email.get('name', '')
    
    # Embed any inline images as data: URIs so they render in place instead of as broken
    # cid: references (mirrors how Connect stores agent-added inline images).
    body_html = email_data.get('body', {}).get('content', '')
    if inline_images:
        body_html = embed_inline_images(body_html, inline_images)
    
    # Build email message structure. MESSAGE_ID uses the real RFC 5322 Message-ID
    # (Graph internetMessageId), not the M365 internal object id.
    email_message = {
        'MessageSourceType': 'RAW',
        'RawMessage': {
            'Subject': email_data.get('subject', 'No Subject'),
            'Body': body_html,
            'ContentType': 'text/html' if email_data.get('body', {}).get('contentType') == 'html' else 'text/plain',
            'Headers': build_raw_headers(email_data)
        }
    }
    
    # Build complete API payload
    payload = {
        'InstanceId': CONNECT_INSTANCE_ID,
        'ContactFlowId': CONNECT_CONTACT_FLOW_ID,
        'FromEmailAddress': {'EmailAddress': sender_email},
        'DestinationEmailAddress': DESTINATION_EMAIL_ADDRESS,
        'EmailMessage': email_message,
        'Attributes': {
            'EmailSubject': email_data.get('subject', '')[:256],
            'EmailFrom': sender_email[:256],
            'AttachmentCount': str(len(attachments))
        }
    }
    
    # Add sender display name if available
    if sender_name:
        payload['FromEmailAddress']['DisplayName'] = sender_name
    
    # Add attachments if present
    if attachments:
        payload['Attachments'] = attachments
    
    # Chain to the prior contact when this email replies to one of our outbound messages
    if related_contact_id:
        payload['RelatedContactId'] = related_contact_id

    # Idempotency: the same M365 message never creates two contacts
    if client_token:
        payload['ClientToken'] = client_token
    
    # Create email contact in Connect
    response = connect_client.start_email_contact(**payload)
    return response['ContactId']

def mark_email_read(token, mailbox_email, message_id):
    """
    Mark an email as read in M365 to prevent reprocessing on the next poll.
    """
    url = _graph_user_url(mailbox_email, 'messages', message_id)
    http_request(url, method='PATCH', 
                data=json.dumps({'isRead': True}),
                headers={'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'})

def lambda_handler(event, context):
    """
    Main Lambda handler - processes unread emails from M365 and creates Connect contacts.
    """
    try:
        print("Starting email processing")

        credentials = load_credentials()
        if secret_not_configured(credentials):
            print("M365 secret not configured, update the Secrets Manager secret")
            return {'statusCode': 200, 'body': json.dumps({'processed': 0})}

        # Authenticate with M365
        token, mailbox_email = get_m365_token(credentials)
        # Our outbound Message-ID domain is the M365 mailbox domain; reuse it to recognise replies
        email_domain = mailbox_email.split('@')[-1] if mailbox_email and '@' in mailbox_email else ''
        connect_client = boto3.client('connect')

        # Fetch unread emails
        emails = get_unread_emails(token, mailbox_email)
        print(f"Found {len(emails)} unread emails")
        
        processed = 0
        for email in emails:
            graph_message_id = email.get('id', '')
            try:
                # Process attachments (and split out inline images for embedding)
                attachments, inline_images = process_email_attachments(token, mailbox_email, graph_message_id)
                
                # Derive RelatedContactId from the reply's In-Reply-To (if it points to one of ours),
                # then confirm the sender is the customer on that contact before chaining.
                in_reply_to = extract_header(email.get('internetMessageHeaders', []), 'In-Reply-To')
                candidate = parse_related_contact_id(in_reply_to, email_domain)
                related_contact_id = None
                if candidate:
                    sender_address = email.get('from', {}).get('emailAddress', {}).get('address', '')
                    related_contact_id = verify_related_contact(connect_client, candidate, sender_address)
                
                # Create email contact in Connect
                contact_id = create_connect_email_contact(
                    email, attachments, related_contact_id, inline_images,
                    client_token=build_client_token(mailbox_email, graph_message_id),
                    connect_client=connect_client)
                print(f"Created Connect contact: {contact_id} graph_message_id={graph_message_id} "
                      f"(RelatedContactId={related_contact_id or 'none'})")
                
                # Mark email as read to prevent reprocessing
                mark_email_read(token, mailbox_email, graph_message_id)
                processed += 1
                
            except Exception as e:
                print(f"Error processing email: graph_message_id={graph_message_id} error={_error_summary(e)}")
        
        print(f"Processing complete - {processed} emails processed")
        return {
            'statusCode': 200,
            'body': json.dumps({'processed': processed})
        }
        
    except Exception as e:
        summary = _error_summary(e)
        print(f"Lambda execution error: {summary}")
        # Re-raise so Lambda records a function error: asynchronous retries and the
        # dead-letter queue only apply to invocations that fail.
        raise RuntimeError(f"inbound poll failed: {summary}") from None
