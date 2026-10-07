"""
Amazon Connect <-> Microsoft 365 outbound email Lambda.

Invoked by the Connect Outbound Whisper flow when an agent sends an email reply.
It assembles an RFC 5322 MIME message and sends it through the Microsoft Graph API,
suppressing Connect's native Amazon SES send with StopContact, while preserving
conversation threading:
  - the outbound Message-ID encodes the Connect ContactId (<contactId@domain>) so the
    customer's reply carries it back in In-Reply-To; and
  - In-Reply-To / References are set to the original inbound Message-ID, resolved from
    the contact's connect:EmailMessageId segment attribute or, failing that, the
    related/initial contact in the chain.
The reply body and attachments are read from Connect-managed storage via
ListContactReferences / GetAttachedFile and ListInstanceStorageConfigs.

Return value (flat string map, consumed by the flow as a STRING_MAP):
  {"status": "sent", "contactId": "<id>"}
  {"status": "failed", "reason": "<short_code>", "contactId": "<id>"}

Required environment variable:
  SECRETS_MANAGER_SECRET_NAME  Secret holding M365 credentials
                               (tenant_id, client_id, client_secret, mailbox_email).

Logging rule: log only ContactIds, operation names, counts, HTTP status codes and Graph
error codes. Never log subjects, filenames, S3 keys, addresses, Internet Message-IDs,
response bodies, draft ids, tokens or secret values.
"""
import json
import boto3
import urllib.request
import urllib.parse
import urllib.error
import os
import base64
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from email import policy
from botocore.exceptions import ClientError

# Environment configuration
SECRETS_NAME = os.environ['SECRETS_MANAGER_SECRET_NAME']

GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
# Short timeout: the whole function must finish inside the flow's Lambda limit. This bounds a
# hung connection, it does not make the normal path faster.
HTTP_TIMEOUT_SECONDS = 10
# Only these hosts (plus *.amazonaws.com, for the GetAttachedFile presigned download) may be
# contacted by urlopen.
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


def _region_from_context(context):
    """Region of this function from its ARN (arn:aws:lambda:REGION:...), else AWS_REGION."""
    try:
        region = context.invoked_function_arn.split(':')[3]
    except (AttributeError, IndexError):
        region = ''
    return region or os.environ['AWS_REGION']


def http_request(url, method='GET', data=None, headers=None):
    """
    HTTP request using urllib - no external dependencies.
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
    """Get an M365 access token (client-credentials flow). Returns (access_token, mailbox_email)."""
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


# ---- Original Message-ID resolution (for threading the reply) ----

def _seg_attr_message_id(contact):
    """Read connect:EmailMessageId.ValueString from a contact/event SegmentAttributes (trimmed)."""
    seg = (contact or {}).get('SegmentAttributes', {}) or {}
    attr = seg.get('connect:EmailMessageId', {}) or {}
    return (attr.get('ValueString') or '').strip()


def _safe_describe(connect_client, instance_id, contact_id, cache):
    """DescribeContact with memoization; failures are non-fatal (return None)."""
    if contact_id in cache:
        return cache[contact_id]
    try:
        resp = connect_client.describe_contact(InstanceId=instance_id, ContactId=contact_id)
        contact = resp.get('Contact', {})
    except Exception as e:
        print(f"DescribeContact failed for {contact_id}: {_error_summary(e)}")
        contact = None
    cache[contact_id] = contact
    return contact


def resolve_original_message_id(connect_client, instance_id, contact_data):
    """
    Resolve the inbound email's Message-ID in order:
      1. outbound contact's own connect:EmailMessageId segment attribute (from the event, no API)
      2. DescribeContact(RelatedContactId)
      3. DescribeContact(InitialContactId)
    At most two DescribeContact calls total; failures are non-fatal; None => send as new thread.
    """
    cache = {}
    budget = 2
    own_id = contact_data.get('ContactId')

    mid = _seg_attr_message_id(contact_data)
    if mid:
        return mid

    related = contact_data.get('RelatedContactId')
    initial = contact_data.get('InitialContactId')

    if related and related != own_id and budget > 0:
        contact = _safe_describe(connect_client, instance_id, related, cache)
        budget -= 1
        mid = _seg_attr_message_id(contact)
        if mid:
            return mid

    if initial and initial != own_id and initial != related and budget > 0:
        contact = _safe_describe(connect_client, instance_id, initial, cache)
        budget -= 1
        mid = _seg_attr_message_id(contact)
        if mid:
            return mid

    return None


# ---- Reply body from Connect attached-file references ----

def get_email_body(connect_client, instance_id, instance_arn, contact_id):
    """Fetch the agent's reply body (and content type) from the EMAIL_MESSAGE attached file."""
    try:
        refs = connect_client.list_contact_references(
            InstanceId=instance_id, ContactId=contact_id, ReferenceTypes=['EMAIL_MESSAGE']
        )
    except Exception as e:
        print(f"ListContactReferences failed: {_error_summary(e)}")
        return None, None

    ref_list = refs.get('ReferenceSummaryList', [])
    if not ref_list:
        print("No EMAIL_MESSAGE reference found")
        return None, None

    file_id = ref_list[0].get('EmailMessage', {}).get('Name')
    if not file_id:
        return None, None

    contact_arn = f"{instance_arn}/contact/{contact_id}"
    try:
        file_resp = connect_client.get_attached_file(
            InstanceId=instance_id, FileId=file_id, AssociatedResourceArn=contact_arn
        )
        download_url = _check_url(file_resp['DownloadUrlMetadata']['Url'])
        with urllib.request.urlopen(urllib.request.Request(download_url), timeout=HTTP_TIMEOUT_SECONDS) as resp:  # nosec B310  # https + *.amazonaws.com host enforced by _check_url
            email_data = json.loads(resp.read())
    except Exception as e:
        print(f"GetAttachedFile/download failed: {_error_summary(e)}")
        return None, None

    return email_data.get('messageContent', ''), email_data.get('contentType')


# ---- Attachments via programmatic storage lookup ----

def resolve_storage(connect_client, instance_id, resource_type):
    """Resolve (bucket, prefix, kms_key_id) for a Connect instance storage config via API."""
    try:
        resp = connect_client.list_instance_storage_configs(
            InstanceId=instance_id, ResourceType=resource_type)
        configs = resp.get('StorageConfigs', [])
        if not configs:
            return None, None, None
        s3 = configs[0].get('S3Config', {})
        kms = (s3.get('EncryptionConfig', {}) or {}).get('KeyId')
        return s3.get('BucketName'), s3.get('BucketPrefix', '') or '', kms
    except Exception as e:
        print(f"ListInstanceStorageConfigs({resource_type}) failed: {_error_summary(e)}")
        return None, None, None


def _find_s3_files(s3_client, bucket, prefix, contact_id):
    """List S3 objects for a contact, scanning the last 7 daily prefixes (newest day first)."""
    if not bucket:
        return []
    current_date = datetime.now(timezone.utc)
    for days_back in range(7):
        search_date = current_date - timedelta(days=days_back)
        date_path = search_date.strftime('%Y/%m/%d')
        search_prefix = f"{prefix}/{date_path}/{contact_id}" if prefix else f"{date_path}/{contact_id}"
        try:
            response = s3_client.list_objects_v2(Bucket=bucket, Prefix=search_prefix)
            if 'Contents' in response:
                return response['Contents']
        except Exception as e:
            print(f"Warning: ListObjectsV2 failed: bucket={bucket} contact={contact_id} "
                  f"days_back={days_back} error={_error_summary(e)}")
            continue
    return []


def get_attachments(connect_client, instance_id, contact_id):
    """Best-effort fetch of outbound attachments from the Connect-managed attachment bucket."""
    bucket, prefix, _kms = resolve_storage(connect_client, instance_id, 'ATTACHMENTS')
    if not bucket:
        return []
    if prefix and not prefix.endswith('/email'):
        prefix = f"{prefix}/email"
    s3_client = boto3.client('s3')
    files = _find_s3_files(s3_client, bucket, prefix, contact_id)

    attachments = []
    for index, obj in enumerate(files):
        try:
            file_response = s3_client.get_object(Bucket=bucket, Key=obj['Key'])
            content = file_response['Body'].read()
            # Skip the email-content JSON if present
            try:
                if b'messageContent' in content and json.loads(content).get('messageContent') is not None:
                    continue
            except Exception:  # nosec B110  # probe only: a non-JSON object is a real attachment, fall through and keep it
                pass
            filename = obj['Key'].split('/')[-1]
            attachments.append({
                'filename': filename,
                'content': content,
                'content_type': file_response.get('ContentType', 'application/octet-stream')
            })
        except Exception as e:
            print(f"Failed to read attachment: contact={contact_id} index={index} error={_error_summary(e)}")
    return attachments


# ---- MIME assembly ----

def build_mime_message(sender, recipient, subject, body, is_html, contact_id, domain,
                       original_message_id=None, attachments=None):
    """
    Build the outbound MIME message. Message-ID encodes the outbound ContactId so the
    customer's reply In-Reply-To carries it back. The domain is the M365 mailbox (From) domain.
    Single body part when no attachments; multipart/mixed when attachments are present. UTF-8 body.
    """
    msg = EmailMessage()
    msg['From'] = sender
    msg['To'] = recipient
    msg['Subject'] = subject
    msg['Message-ID'] = f"<{contact_id}@{domain}>"
    if original_message_id:
        msg['In-Reply-To'] = original_message_id
        msg['References'] = original_message_id

    if is_html:
        msg.set_content(body or '', subtype='html', charset='utf-8')
    else:
        msg.set_content(body or '', charset='utf-8')

    for att in attachments or []:
        ctype = att.get('content_type') or 'application/octet-stream'
        maintype, _, subtype = ctype.partition('/')
        if not subtype:
            maintype, subtype = 'application', 'octet-stream'
        msg.add_attachment(att['content'], maintype=maintype, subtype=subtype,
                           filename=att.get('filename', 'attachment'))
    return msg


def delete_draft(token, mailbox, draft_id):
    """Delete a Graph draft message (cleanup after a failed send). No-op if draft_id is falsy."""
    if not draft_id:
        return
    try:
        url = _graph_user_url(mailbox, 'messages', draft_id)
        http_request(url, method='DELETE', headers={'Authorization': f'Bearer {token}'})
        print("Deleted orphaned draft")
    except Exception as e:
        print(f"Failed to delete draft: {_error_summary(e)}")


def send_mime_via_graph(token, mailbox, mime_bytes):
    """Create a MIME draft then send it, preserving custom Message-ID/In-Reply-To/References."""
    encoded = base64.b64encode(mime_bytes).decode()
    create_url = _graph_user_url(mailbox, 'messages')
    resp = http_request(create_url, method='POST', data=encoded,
                        headers={'Authorization': f'Bearer {token}', 'Content-Type': 'text/plain'})
    draft = json.loads(resp['content'])
    draft_id = draft.get('id')

    try:
        send_url = _graph_user_url(mailbox, 'messages', draft_id, 'send')
        http_request(send_url, method='POST',
                     headers={'Authorization': f'Bearer {token}', 'Content-Length': '0'})
        return draft_id
    except Exception as e:
        print(f"Graph send failed; cleaning up draft: {_error_summary(e)}")
        delete_draft(token, mailbox, draft_id)
        raise


def stop_contact(contact_id, instance_id):
    """Stop the Connect contact to prevent the native SES send (API-only bypass)."""
    try:
        connect_client = boto3.client('connect')
        connect_client.stop_contact(ContactId=contact_id, InstanceId=instance_id)
        print(f"Contact {contact_id} stopped successfully")
    except Exception as e:
        print(f"Warning: StopContact failed for {contact_id}: {_error_summary(e)}")


def _result(status, contact_id, reason=None):
    """Flat string map for the flow (STRING_MAP): status, contactId and, on failure, a short reason code."""
    result = {'status': status, 'contactId': str(contact_id or 'unknown')}
    if reason:
        result['reason'] = str(reason)
    return result


class OutboundError(Exception):
    """Failure with a short reason code that is safe to return to the flow and to log."""

    def __init__(self, reason):
        self.reason = reason
        super().__init__(reason)


def lambda_handler(event, context):
    """Outbound handler invoked from the Connect Outbound Whisper flow."""
    contact_id = None
    instance_id = None
    reason = 'internal_error'
    try:
        contact_data = event.get('Details', {}).get('ContactData', {})
        contact_id = contact_data.get('ContactId')
        recipient = contact_data.get('CustomerEndpoint', {}).get('Address')
        subject = contact_data.get('SegmentAttributes', {}).get(
            'connect:EmailSubject', {}).get('ValueString', 'Re: Email')

        instance_arn = contact_data.get('InstanceARN', '')
        instance_id = contact_data.get('Tags', {}).get('aws:connect:instanceId')
        if not instance_id and instance_arn:
            instance_id = instance_arn.split('/')[-1]
        if not instance_arn and instance_id:
            region = _region_from_context(context)
            account = context.invoked_function_arn.split(':')[4]
            instance_arn = f"arn:aws:connect:{region}:{account}:instance/{instance_id}"

        if not contact_id or not recipient or not instance_id:
            print(f"Missing required data: contact_id={contact_id or 'none'} "
                  f"instance_id={instance_id or 'none'} has_recipient={bool(recipient)}")
            raise OutboundError('missing_contact_data')

        connect_client = boto3.client('connect')

        reason = 'secret_unavailable'
        credentials = load_credentials()
        if secret_not_configured(credentials):
            print("M365 secret not configured, update the Secrets Manager secret")
            raise OutboundError('secret_not_configured')

        reason = 'token_failed'
        token, sender = get_m365_token(credentials)
        # Outbound Message-ID domain = the M365 mailbox (From) domain
        email_domain = sender.split('@')[-1] if sender and '@' in sender else ''

        reason = 'prepare_failed'
        original_message_id = resolve_original_message_id(connect_client, instance_id, contact_data)
        body, content_type = get_email_body(connect_client, instance_id, instance_arn, contact_id)
        if not body:
            body, content_type = "Thank you for your email.", content_type or 'text/html'
        is_html = (content_type or '').lower() in ('html', 'text/html')

        attachments = get_attachments(connect_client, instance_id, contact_id)

        mime = build_mime_message(sender, recipient, subject, body, is_html, contact_id, email_domain,
                                  original_message_id, attachments)

        reason = 'graph_send_failed'
        # Serialize with the SMTP policy so MIME uses CRLF line endings.
        # Quoted-printable soft breaks must be "=\r\n"; bare-LF "=\n" gets
        # mangled by Exchange into stray "=" characters in the delivered body.
        send_mime_via_graph(token, sender, mime.as_bytes(policy=policy.SMTP))
        print(f"Sent reply: contact={contact_id}, threaded={bool(original_message_id)}, "
              f"attachments={len(attachments)}")

        # Stop the contact so Connect/SES does not also send (success path)
        stop_contact(contact_id, instance_id)
        return _result('sent', contact_id)

    except Exception as e:
        if isinstance(e, OutboundError):
            reason = e.reason
        print(f"Error: reason={reason} contact={contact_id or 'none'} error={_error_summary(e)}")
        # Stop the contact even on failure to prevent a duplicate Connect/SES send
        if contact_id and instance_id:
            stop_contact(contact_id, instance_id)
        return _result('failed', contact_id, reason)
