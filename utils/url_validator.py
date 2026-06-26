import re
import os
import requests
import logging


def _ssl_verify():
    if os.environ.get('DISABLE_SSL_VERIFY', '').lower() in ('1', 'true', 'yes'):
        return False
    return os.environ.get('REQUESTS_CA_BUNDLE') or os.environ.get('SSL_CERT_FILE') or True

logger = logging.getLogger(__name__)

# Regex to find URLs in text
URL_PATTERN = re.compile(
    r"""(?i)\b((?:https?://|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:'".,<>?«»“”‘’]))""", re.IGNORECASE
)

def validate_urls_in_text(text: str) -> list[dict]:
    """
    Extracts URLs from text and validates their reachability.
    Returns a list of dictionaries with URL, status, and status_code.
    """
    urls_found = URL_PATTERN.findall(text)
    results = []

    for url_match in urls_found:
        # url_match is a tuple, the first element is the full match
        url = url_match[0]
        status = "unknown"
        status_code = None

        try:
            # Use HEAD request to minimize bandwidth
            response = requests.head(url, timeout=5, allow_redirects=True, verify=_ssl_verify())
            status_code = response.status_code

            if 200 <= status_code < 400:
                status = "valid"
            elif status_code == 404:
                status = "broken"
            else:
                status = "unreachable"
        except requests.exceptions.RequestException as e:
            status = "unreachable"
            logger.warning(f"Could not validate URL {url}: {e}")

        results.append({"url": url, "status": status, "status_code": status_code})

    return results
