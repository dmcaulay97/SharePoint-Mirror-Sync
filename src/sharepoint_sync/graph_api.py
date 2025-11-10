# -*- coding: utf-8 -*-
"""
Microsoft Graph API operations for SharePoint sync.

This module provides all Graph API interactions including column management,
list item operations, and request retry logic.
"""

import time
import requests
from dotenv import load_dotenv
from .auth import acquire_token
from .monitoring import rate_monitor
from .utils import is_debug_metadata_enabled, is_debug_enabled

# Load environment variables
load_dotenv()

# Global cache for column mappings
column_mapping_cache = {}

# Global cache for site/drive IDs (used by deletion operations)
site_drive_id_cache = {}


def make_graph_request_with_retry(url, headers, method='GET', json_data=None, data=None, params=None, max_retries=3):
    """
    Make a Graph API request with proper retry handling for transient errors.
    Includes rate limiting monitoring via response header analysis.

    Retry Logic:
        - 429 (Rate Limit): Waits for Retry-After header duration
        - 5xx (Server Error): Exponential backoff (1s, 3s, 7s)
        - 409 (Conflict/Lock): Exponential backoff (2s, 4s, 8s) - files being processed
        - 4xx (Client Error): No retry (except 409)

    Args:
        url (str): The Graph API endpoint URL
        headers (dict): Request headers including Authorization
        method (str): HTTP method ('GET', 'POST', 'PATCH', 'PUT', 'DELETE', etc.)
        json_data (dict): JSON data for POST/PATCH requests (mutually exclusive with data)
        data (bytes): Binary data for PUT/POST requests (mutually exclusive with json_data)
        params (dict): URL parameters for GET requests
        max_retries (int): Maximum number of retry attempts (default: 3)

    Returns:
        requests.Response: The HTTP response object

    Raises:
        Exception: If all retries are exhausted for 429 or 5xx errors

    Note:
        Use json_data for JSON requests or data for binary uploads, not both.
        409 errors return response after retries (no exception) for graceful handling.
    """
    debug_metadata = is_debug_metadata_enabled()

    for attempt in range(max_retries + 1):
        try:
            # Add proactive delay if approaching rate limits
            if rate_monitor.should_slow_down() and attempt > 0:
                delay = 2 ** attempt
                if is_debug_enabled():
                    print(f"[⚠] Proactive rate limiting delay: {delay}s")
                time.sleep(delay)

            # Make the request based on method
            if method.upper() == 'GET':
                response = requests.get(url, headers=headers, params=params)
            elif method.upper() == 'POST':
                if data is not None:
                    response = requests.post(url, headers=headers, data=data)
                else:
                    response = requests.post(url, headers=headers, json=json_data)
            elif method.upper() == 'PATCH':
                response = requests.patch(url, headers=headers, json=json_data)
            elif method.upper() == 'PUT':
                if data is not None:
                    response = requests.put(url, headers=headers, data=data)
                else:
                    response = requests.put(url, headers=headers, json=json_data)
            elif method.upper() == 'DELETE':
                response = requests.delete(url, headers=headers)
            else:
                raise ValueError(f"Unsupported HTTP method: {method}")

            # Analyze response headers for rate limiting info (with request type tracking)
            rate_monitor.analyze_response_headers(response, method=method, url=url)

            # Check for rate limiting (429) or server errors (5xx)
            if response.status_code == 429:
                # Get retry-after header value
                retry_after = response.headers.get('Retry-After', '60')
                try:
                    wait_seconds = int(retry_after)
                except ValueError:
                    wait_seconds = 60  # Default to 60 seconds if header is malformed

                if attempt < max_retries:
                    if is_debug_enabled():
                        print(f"[!] Rate limited (429). Waiting {wait_seconds} seconds before retry {attempt + 1}/{max_retries}...")
                    if debug_metadata:
                        print(f"[DEBUG] Retry-After header: {retry_after}")
                        print(f"[DEBUG] Rate limit response: {response.text[:300]}")
                    time.sleep(wait_seconds)
                    continue
                else:
                    print(f"[!] Rate limiting exhausted all retries. Final 429 response:")
                    print(f"[DEBUG] {response.text[:500]}")
                    raise Exception(f"Graph API rate limiting: {response.status_code} after {max_retries} retries")

            elif 500 <= response.status_code < 600:
                # Server error - retry with exponential backoff
                if attempt < max_retries:
                    wait_seconds = (2 ** attempt) + 1  # 1, 3, 7 seconds
                    if is_debug_enabled():
                        print(f"[!] Server error ({response.status_code}). Retrying in {wait_seconds} seconds... ({attempt + 1}/{max_retries})")
                    if debug_metadata:
                        print(f"[DEBUG] Server error response: {response.text[:300]}")
                    time.sleep(wait_seconds)
                    continue
                else:
                    if is_debug_enabled():
                        print(f"[!] Server errors exhausted all retries. Final response:")
                    print(f"[DEBUG] {response.text[:500]}")
                    raise Exception(f"Graph API server error: {response.status_code} after {max_retries} retries")

            elif response.status_code == 409:
                # Conflict error (file locked, being processed, etc.) - retry with exponential backoff
                # This is often transient (SharePoint processing, virus scan, indexing)
                if attempt < max_retries:
                    wait_seconds = (2 ** attempt) + 2  # 2, 4, 8 seconds (longer than server errors)
                    if is_debug_enabled():
                        print(f"[!] Conflict/Lock error (409). File may be locked or processing. Retrying in {wait_seconds} seconds... ({attempt + 1}/{max_retries})")
                    if debug_metadata:
                        print(f"[DEBUG] Conflict response: {response.text[:300]}")
                    time.sleep(wait_seconds)
                    continue
                else:
                    if is_debug_enabled():
                        print(f"[!] Conflict errors exhausted all retries. File may be locked.")
                    if debug_metadata:
                        print(f"[DEBUG] Final 409 response: {response.text[:500]}")
                    # Don't raise exception - return response to allow graceful handling
                    return response

            # Success or client error (don't retry client errors like 400, 401, 403, 404)
            return response

        except requests.exceptions.Timeout as e:
            # Request timeout - retry with exponential backoff
            if attempt < max_retries:
                wait_seconds = (2 ** attempt) + 1
                timeout_info = str(e)[:100] if str(e) else "timeout"
                print(f"[!] Request timeout ({timeout_info}). Retrying in {wait_seconds} seconds... ({attempt + 1}/{max_retries})")
                time.sleep(wait_seconds)
                continue
            else:
                print("[!] ========================================")
                print("[!] REQUEST TIMEOUT - All retries exhausted")
                print("[!] ========================================")
                print(f"[!] The request to Graph API timed out after {max_retries} retry attempts.")
                print("[!] ")
                print("[!] Troubleshooting steps:")
                print("[!]   1. Check your internet connection speed")
                print("[!]   2. Verify network connectivity to Microsoft Graph API")
                print("[!]   3. If using a proxy, verify proxy configuration")
                print("[!]   4. Try again - Microsoft services may be experiencing issues")
                print("[!]   5. For large file uploads, this may indicate a very slow connection")
                print("[!] ")
                print(f"[!] URL: {url[:100]}...")
                print("[!] ========================================")
                raise Exception(f"Graph API request timed out after {max_retries} retries. Check network connectivity.")

        except requests.exceptions.SSLError as e:
            # SSL errors usually aren't transient - fail fast with clear message
            print("[!] ========================================")
            print("[!] SSL/TLS CERTIFICATE ERROR")
            print("[!] ========================================")
            print("[!] Failed to verify SSL certificate for Microsoft Graph API.")
            print("[!] ")
            print("[!] Troubleshooting steps:")
            print("[!]   1. Verify system certificate store is up to date")
            print("[!]   2. Check if corporate proxy is intercepting SSL/TLS connections")
            print("[!]   3. Ensure system clock is accurate (SSL cert validation requires correct time)")
            print("[!]   4. Try updating Python's certifi package: pip install --upgrade certifi")
            print("[!]   5. If behind a corporate firewall, you may need to import company's root CA")
            print("[!] ")
            print(f"[!] Technical details: {str(e)[:300]}")
            print(f"[!] URL: {url[:100]}...")
            print("[!] ========================================")
            raise Exception(f"SSL certificate verification failed: {str(e)[:200]}")

        except requests.exceptions.ProxyError as e:
            # Proxy connection errors - fail fast with configuration guidance
            print("[!] ========================================")
            print("[!] PROXY CONNECTION ERROR")
            print("[!] ========================================")
            print("[!] Failed to connect through proxy server.")
            print("[!] ")
            print("[!] Troubleshooting steps:")
            print("[!]   1. Verify HTTP_PROXY and HTTPS_PROXY environment variables are set correctly")
            print("[!]   2. Check proxy server is accessible and responding")
            print("[!]   3. Verify proxy authentication credentials if required")
            print("[!]   4. Test direct connection (temporarily disable proxy) to isolate issue")
            print("[!]   5. Check proxy server allows connections to *.microsoft.com")
            print("[!] ")
            print(f"[!] Technical details: {str(e)[:300]}")
            print("[!] ========================================")
            raise Exception(f"Proxy connection failed: {str(e)[:200]}")

        except requests.exceptions.TooManyRedirects as e:
            # Redirect loop - indicates configuration issue
            print("[!] ========================================")
            print("[!] TOO MANY REDIRECTS")
            print("[!] ========================================")
            print("[!] Encountered redirect loop - this indicates a configuration issue.")
            print("[!] ")
            print("[!] Troubleshooting steps:")
            print(f"[!]   1. Verify Graph API endpoint is correct: {url[:100]}...")
            print("[!]   2. Check if proxy is misconfigured and causing redirect loops")
            print("[!]   3. Verify you're using the correct cloud endpoint:")
            print("[!]      - Commercial: graph.microsoft.com")
            print("[!]      - GovCloud: graph.microsoft.us")
            print("[!] ")
            print(f"[!] Technical details: {str(e)[:300]}")
            print("[!] ========================================")
            raise Exception(f"Too many redirects - possible configuration issue: {str(e)[:200]}")

        except requests.exceptions.ConnectionError as e:
            # Network/DNS connection errors - retry with exponential backoff
            if attempt < max_retries:
                wait_seconds = (2 ** attempt) + 1
                error_detail = str(e)[:100]
                print(f"[!] Network connection error: {error_detail}. Retrying in {wait_seconds} seconds... ({attempt + 1}/{max_retries})")
                time.sleep(wait_seconds)
                continue
            else:
                print("[!] ========================================")
                print("[!] NETWORK CONNECTION FAILED")
                print("[!] ========================================")
                print(f"[!] Could not establish connection after {max_retries} retry attempts.")
                print("[!] ")
                print("[!] Troubleshooting steps:")
                print("[!]   1. Verify internet connectivity (try: ping 8.8.8.8)")
                print("[!]   2. Check DNS resolution (try: nslookup graph.microsoft.com)")
                print("[!]   3. Ensure firewall allows HTTPS (port 443) to *.microsoft.com")
                print("[!]   4. If using VPN, verify VPN connection is stable")
                print("[!]   5. Try disabling any VPN/proxy temporarily to isolate issue")
                print("[!]   6. Check Microsoft Azure status page for service outages")
                print("[!] ")
                print(f"[!] Technical details: {str(e)[:300]}")
                print("[!] ========================================")
                raise Exception(f"Network connection failed after {max_retries} retries: {str(e)[:200]}")

        except requests.exceptions.RequestException as e:
            # Catch-all for other request errors (should be rare after specific catches above)
            if attempt < max_retries:
                wait_seconds = (2 ** attempt) + 1
                print(f"[!] HTTP request error: {str(e)[:100]}. Retrying in {wait_seconds} seconds... ({attempt + 1}/{max_retries})")
                time.sleep(wait_seconds)
                continue
            else:
                print(f"[!] HTTP request errors exhausted all retries: {str(e)[:200]}")
                raise

    # Should never reach here, but just in case
    raise Exception("Unexpected error in make_graph_request_with_retry")


def get_column_internal_name_mapping(site_id, list_id, token, graph_endpoint):
    """
    Get mapping of display names to internal names for all columns in a SharePoint list.

    Args:
        site_id (str): SharePoint site ID
        list_id (str): SharePoint list/library ID
        token (str): OAuth access token
        graph_endpoint (str): Microsoft Graph API endpoint

    Returns:
        dict: Mapping of display names to column metadata including internal names
              Format: {display_name: {'internal_name': str, 'type': str, 'id': str, 'description': str}}

    Note:
        Results are cached globally in column_mapping_cache to reduce API calls.
    """
    global column_mapping_cache

    # Check cache first
    cache_key = (site_id, list_id)
    if cache_key in column_mapping_cache:
        debug_metadata = is_debug_metadata_enabled()
        if debug_metadata:
            print(f"[=] Using cached column mappings for site/list")
        return column_mapping_cache[cache_key]

    try:
        debug_metadata = is_debug_metadata_enabled()

        url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists/{list_id}/columns"
        headers = {
            'Authorization': f'Bearer {token}',
            'Accept': 'application/json',
            'Content-Type': 'application/json'
        }

        if debug_metadata:
            print(f"[=] Fetching column mappings from Graph API...")

        response = make_graph_request_with_retry(url, headers, method='GET')

        if response.status_code == 200:
            columns = response.json().get('value', [])
            mapping = {}

            for column in columns:
                display_name = column.get('displayName', '')
                internal_name = column.get('name', '')
                column_type = column.get('columnGroup', 'Unknown')

                mapping[display_name] = {
                    'internal_name': internal_name,
                    'type': column_type,
                    'id': column.get('id', ''),
                    'description': column.get('description', '')
                }

                if debug_metadata:
                    if is_debug_enabled():
                        print(f"[=] Column mapping: '{display_name}' -> '{internal_name}' ({column_type})")

            # Cache the result
            column_mapping_cache[cache_key] = mapping

            if debug_metadata:
                if is_debug_enabled():
                    print(f"[OK] Cached {len(mapping)} column mappings")

            return mapping
        else:
            print(f"[!] Failed to get column mapping: {response.status_code}")
            if debug_metadata:
                print(f"[DEBUG] Response: {response.text[:500]}")
            return {}

    except Exception as e:
        print(f"[!] Error getting column mapping: {e}")
        return {}


def resolve_field_name(site_id, list_id, token, graph_endpoint, field_name):
    """
    Resolve display name to internal name for reliable field access.

    SharePoint columns have both display names (what users see) and internal names
    (used by API). This function resolves display names to their internal counterparts.

    Args:
        site_id (str): SharePoint site ID
        list_id (str): SharePoint list/library ID
        token (str): OAuth access token
        graph_endpoint (str): Microsoft Graph API endpoint
        field_name (str): Display name or internal name to resolve

    Returns:
        str: The internal name for the field, or original name if not resolved

    Note:
        - Internal names use hex codes for special characters (e.g., '_x0020_' for space)
        - If field_name already appears to be an internal name, returns it as-is
        - Falls back to case-insensitive matching if exact match not found
    """
    try:
        debug_metadata = is_debug_metadata_enabled()

        # First check if it's already an internal name by checking for hex encoding
        if '_x00' in field_name or (not any(c.isupper() for c in field_name) and '_' in field_name):
            if debug_metadata:
                if is_debug_enabled():
                    print(f"[=] '{field_name}' appears to be internal name (contains hex encoding)")
            return field_name

        # Get column mapping
        column_mapping = get_column_internal_name_mapping(site_id, list_id, token, graph_endpoint)

        # Try exact display name match
        if field_name in column_mapping:
            internal_name = column_mapping[field_name]['internal_name']
            if debug_metadata:
                if is_debug_enabled():
                    print(f"[OK] Resolved '{field_name}' to internal name '{internal_name}'")
            return internal_name

        # Try case-insensitive match
        for display_name, details in column_mapping.items():
            if display_name.lower() == field_name.lower():
                internal_name = details['internal_name']
                if debug_metadata:
                    print(f"[OK] Resolved '{field_name}' to internal name '{internal_name}' (case-insensitive)")
                return internal_name

        # If no match found, return original name
        if debug_metadata:
            print(f"[!] Could not resolve '{field_name}' to internal name, using as-is")
        return field_name

    except Exception as e:
        print(f"[!] Error resolving field name: {e}")
        return field_name


def sanitize_field_name_for_sharepoint(field_name):
    """
    Convert display name to expected internal name format by encoding special characters.

    SharePoint internal names encode special characters as hex values (e.g., '_x0020_' for space).
    This function attempts to convert a display name to its likely internal name format.

    Args:
        field_name (str): Display name to sanitize

    Returns:
        str: Sanitized field name with special characters encoded

    Note:
        This is a fallback mechanism. Prefer using resolve_field_name() with Graph API
        for accurate internal name resolution.

    Examples:
        'File Hash' -> 'File_x0020_Hash'
        'User#ID' -> 'User_x0023_ID'
        'Value%' -> 'Value_x0025_'
    """
    # Handle common special character conversions
    replacements = {
        ' ': '_x0020_',
        '#': '_x0023_',
        '%': '_x0025_',
        '&': '_x0026_',
        '*': '_x002a_',
        '+': '_x002b_',
        '/': '_x002f_',
        ':': '_x003a_',
        '<': '_x003c_',
        '>': '_x003e_',
        '?': '_x003f_',
        '\\': '_x005c_',
        '|': '_x007c_'
    }

    sanitized = field_name
    for char, replacement in replacements.items():
        sanitized = sanitized.replace(char, replacement)

    return sanitized


def check_and_create_filehash_column(site_url, list_name, tenant_id, client_id, client_secret, login_endpoint, graph_endpoint):
    """
    Check if FileHash column exists in SharePoint document library and create if needed.

    Uses direct Graph API calls to bypass Office365-REST-Python-Client limitations.
    This ensures the FileHash column is available for storing file hashes.

    Args:
        site_url (str): Full SharePoint site URL
        list_name (str): Name of the document library (usually "Documents")
        tenant_id (str): Azure AD tenant ID
        client_id (str): App registration client ID
        client_secret (str): App registration client secret
        login_endpoint (str): Azure AD endpoint
        graph_endpoint (str): Graph API endpoint

    Returns:
        tuple: (success: bool, actual_library_name: str)
               - success: True if column exists or was created successfully
               - actual_library_name: The library name that was actually used (may be fallback)

    Note:
        Requires Sites.ReadWrite.All or Sites.Manage.All permissions.
        The column is created as a single line of text with 255 character limit
        (exact length of xxHash128 hexadecimal representation).
    """
    try:
        # Get token for Graph API
        print("[?] Checking for FileHash column in SharePoint...")
        token = acquire_token(tenant_id, client_id, client_secret, login_endpoint, graph_endpoint)

        if 'access_token' not in token:
            print(f"[!] Failed to acquire token for Graph API: {token.get('error_description', 'Unknown error')}")
            return False, list_name

        headers = {
            'Authorization': f"Bearer {token['access_token']}",
            'Content-Type': 'application/json'
        }

        # Parse site URL to get site ID
        # Format: https://tenant.sharepoint.com/sites/sitename
        site_parts = site_url.replace('https://', '').split('/')
        host_name = site_parts[0]
        site_name = site_parts[2] if len(site_parts) > 2 else ''

        # Get site ID first
        site_endpoint = f"https://{graph_endpoint}/v1.0/sites/{host_name}:/sites/{site_name}"
        site_response = make_graph_request_with_retry(site_endpoint, headers, method='GET')

        if site_response.status_code != 200:
            print(f"[!] Failed to get site information: {site_response.status_code}")
            print(f"[DEBUG] Response: {site_response.text[:500]}")
            return False, list_name

        site_data = site_response.json()
        site_id = site_data.get('id')

        if not site_id:
            print("[!] Could not retrieve site ID")
            return False, list_name

        # Get the document library (list) ID
        lists_endpoint = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists"
        lists_response = make_graph_request_with_retry(lists_endpoint, headers, method='GET')

        if lists_response.status_code != 200:
            print(f"[!] Failed to get lists: {lists_response.status_code}")
            return False, list_name

        lists_data = lists_response.json()
        list_id = None
        actual_library_name = list_name

        # Find the document library by name
        for lst in lists_data.get('value', []):
            if lst.get('displayName') == list_name or lst.get('name') == list_name:
                list_id = lst.get('id')
                break

        if not list_id:
            # Try "Shared Documents" as fallback
            for lst in lists_data.get('value', []):
                if lst.get('displayName') == 'Shared Documents' or lst.get('name') == 'Shared Documents':
                    list_id = lst.get('id')
                    actual_library_name = 'Shared Documents'
                    if is_debug_enabled():
                        print(f"[DEBUG] Using 'Shared Documents' instead of '{list_name}'")
                    break

        if not list_id:
            print(f"[!] Document library '{list_name}' not found")
            return False, list_name

        # Check if FileHash column already exists
        columns_endpoint = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists/{list_id}/columns"
        columns_response = make_graph_request_with_retry(columns_endpoint, headers, method='GET')

        if columns_response.status_code != 200:
            print(f"[!] Failed to get columns: {columns_response.status_code}")
            return False, actual_library_name

        columns_data = columns_response.json()
        filehash_exists = False

        # Check for existing FileHash column
        for column in columns_data.get('value', []):
            if column.get('name') == 'FileHash' or column.get('displayName') == 'FileHash':
                filehash_exists = True
                print("[✓] FileHash column already exists")
                break

        # Create column if it doesn't exist
        if not filehash_exists:
            print("[+] Creating FileHash column...")

            # Column definition for FileHash
            column_definition = {
                "displayName": "FileHash",
                "name": "FileHash",
                "description": "xxHash128 checksum for file content verification",
                "enforceUniqueValues": False,
                "hidden": False,
                "indexed": False,
                "readOnly": False,
                "required": False,
                "text": {
                    "allowMultipleLines": False,
                    "appendChangesToExistingText": False,
                    "linesForEditing": 0,
                    "maxLength": 255  # xxHash128 produces 32-character hex string
                }
            }

            # Create the column with retry handling
            create_response = make_graph_request_with_retry(
                columns_endpoint,
                headers,
                method='POST',
                json_data=column_definition
            )

            if create_response.status_code == 201:
                print("[✓] FileHash column created successfully")
                # Wait briefly for column to be fully available (eventual consistency)
                time.sleep(2)

                # Verify the newly created column
                is_valid, validation_msg = verify_column_for_filehash_operations(
                    site_id, list_id, token['access_token'], graph_endpoint
                )
                if not is_valid:
                    print(f"[⚠] FileHash column created but verification failed: {validation_msg}")
                    print(f"[⚠] Column may not be immediately accessible (eventual consistency)")
                    # Still return True since column was created, just not immediately accessible

                return True, actual_library_name
            else:
                print(f"[!] Failed to create FileHash column: {create_response.status_code}")
                print(f"[DEBUG] Response: {create_response.text[:500]}")
                return False, actual_library_name

        # Column already exists - verify it's suitable for operations
        is_valid, validation_msg = verify_column_for_filehash_operations(
            site_id, list_id, token['access_token'], graph_endpoint
        )

        if not is_valid:
            print(f"[⚠] FileHash column exists but has issues: {validation_msg}")
            print(f"[⚠] Hash-based comparison may not work correctly")
            # Still return True since column exists, but warn about issues

        return True, actual_library_name

    except Exception as e:
        print(f"[!] Error checking/creating FileHash column: {e}")
        return False, list_name


def rewrite_endpoint(request, graph_endpoint):
    """
    Modify API request URLs for non-standard Microsoft Graph endpoints.

    This function is needed for special Azure environments like:
    - Azure Government Cloud (graph.microsoft.us)
    - Azure Germany (graph.microsoft.de)
    - Azure China (microsoftgraph.chinacloudapi.cn)

    Args:
        request: The HTTP request object to be modified
        graph_endpoint (str): The target Graph API endpoint

    Note:
        This is a callback function used by the GraphClient to intercept
        and modify requests before they're sent.
    """
    # Replace default endpoint with custom one if specified
    request.url = request.url.replace(
        "https://graph.microsoft.com", f"https://{graph_endpoint}"
    )


def update_sharepoint_list_item_field(site_url, list_name, item_id, field_name, field_value, tenant_id, client_id, client_secret, login_endpoint, graph_endpoint):
    """
    Update a custom field in a SharePoint list item using direct Graph API REST calls.

    Args:
        site_url (str): Full SharePoint site URL
        list_name (str): Name of the document library (usually "Documents")
        item_id (str): SharePoint list item ID
        field_name (str): Internal name of the field to update
        field_value (str): Value to set for the field
        tenant_id (str): Azure AD tenant ID
        client_id (str): App registration client ID
        client_secret (str): App registration client secret
        login_endpoint (str): Azure AD endpoint
        graph_endpoint (str): Graph API endpoint

    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Get debug flag
        debug_metadata = is_debug_metadata_enabled()

        # Get token for Graph API
        token = acquire_token(tenant_id, client_id, client_secret, login_endpoint, graph_endpoint)

        if 'access_token' not in token:
            print(f"[!] Failed to acquire token for Graph API: {token.get('error_description', 'Unknown error')}")
            return False

        headers = {
            'Authorization': f"Bearer {token['access_token']}",
            'Content-Type': 'application/json'
        }

        # Check for rate limiting headers
        if debug_metadata:
            print(f"[DEBUG] Updating field {field_name} = {field_value} for item {item_id}")

        # Parse site URL to get site ID
        site_parts = site_url.replace('https://', '').split('/')
        host_name = site_parts[0]
        site_name = site_parts[2] if len(site_parts) > 2 else ''

        # Get site ID first
        site_endpoint = f"https://{graph_endpoint}/v1.0/sites/{host_name}:/sites/{site_name}"
        site_response = make_graph_request_with_retry(site_endpoint, headers=headers)

        if site_response.status_code != 200:
            print(f"[!] Failed to get site information: {site_response.status_code}")
            if debug_metadata:
                print(f"[DEBUG] Site response: {site_response.text[:300]}")
            return False

        site_data = site_response.json()
        site_id = site_data.get('id')

        if not site_id:
            print("[!] Could not retrieve site ID")
            return False

        # Get the document library (list) ID
        lists_endpoint = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists"
        lists_response = make_graph_request_with_retry(lists_endpoint, headers=headers)

        if lists_response.status_code != 200:
            print(f"[!] Failed to get lists: {lists_response.status_code}")
            if debug_metadata:
                print(f"[DEBUG] Lists response: {lists_response.text[:300]}")
            return False

        lists_data = lists_response.json()
        list_id = None

        for sp_list in lists_data.get('value', []):
            if sp_list.get('displayName') == list_name or sp_list.get('name') == list_name:
                list_id = sp_list.get('id')
                break

        if not list_id:
            print(f"[!] Could not find list '{list_name}'")
            return False

        # Resolve field name to internal name for reliable API access
        resolved_field_name = resolve_field_name(site_id, list_id, token['access_token'], graph_endpoint, field_name)

        if resolved_field_name != field_name and debug_metadata:
            print(f"[=] Resolved field name '{field_name}' to '{resolved_field_name}'")

        # Update the field using PATCH request
        fields_endpoint = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists/{list_id}/items/{item_id}/fields"
        field_data = {resolved_field_name: field_value}

        if debug_metadata:
            print(f"[DEBUG] PATCH endpoint: {fields_endpoint}")
            print(f"[DEBUG] Field data to update: {field_data}")

        update_response = requests.patch(fields_endpoint, headers=headers, json=field_data)

        # Check for rate limiting headers in response
        if debug_metadata:
            rate_limit_headers = {}
            for header_name, header_value in update_response.headers.items():
                if 'rate' in header_name.lower() or 'throttl' in header_name.lower() or 'limit' in header_name.lower():
                    rate_limit_headers[header_name] = header_value
            if rate_limit_headers:
                print(f"[DEBUG] Rate limiting headers: {rate_limit_headers}")

        if update_response.status_code == 200:
            if debug_metadata:
                print(f"[DEBUG] ✓ Field update successful")
                # Show updated field data
                response_data = update_response.json()
                if field_name in response_data:
                    print(f"[DEBUG] Confirmed field value: {response_data[field_name]}")
            return True
        elif update_response.status_code == 429:
            # Handle throttling specifically
            retry_after = update_response.headers.get('Retry-After', '60')
            print(f"[!] Request throttled (429). Should wait {retry_after} seconds before retry")
            print(f"[DEBUG] Throttling response: {update_response.text[:500]}")
            return False
        else:
            print(f"[!] Failed to update field: {update_response.status_code}")
            print(f"[DEBUG] Response: {update_response.text[:500]}")

            if debug_metadata:
                print(f"[DEBUG] Request headers: {dict(headers)}")
                print(f"[DEBUG] Response headers: {dict(update_response.headers)}")

                # Check if the field name exists
                if update_response.status_code == 400:
                    print(f"[DEBUG] Bad request - field '{field_name}' may not exist or have wrong internal name")

            return False

    except Exception as e:
        print(f"[!] Error updating list item field: {str(e)[:400]}")
        if is_debug_metadata_enabled():
            import traceback
            print(f"[DEBUG] Full traceback: {traceback.format_exc()}")
        return False


def test_column_accessibility(site_id, list_id, token, graph_endpoint, internal_name):
    """
    Test if a column is accessible by trying to read from list items.

    This function performs a selective query to verify that a column can be accessed
    and read from the SharePoint list. This is useful for detecting columns that exist
    but are not available due to permissions or other restrictions.

    Args:
        site_id (str): SharePoint site ID
        list_id (str): SharePoint list/library ID
        token (str): OAuth access token
        graph_endpoint (str): Microsoft Graph API endpoint
        internal_name (str): Internal name of the column to test

    Returns:
        bool: True if column is accessible, False otherwise

    Note:
        Uses $select to request specific field, which will fail if field is not accessible
    """
    try:
        debug_metadata = is_debug_metadata_enabled()

        # Try to get list items with specific field selection
        # Must explicitly request custom columns in $expand for them to appear
        url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists/{list_id}/items"
        params = {
            '$top': 1,
            '$expand': f'fields($select={internal_name})',
            '$select': f'id,fields'
        }
        headers = {
            'Authorization': f'Bearer {token}',
            'Accept': 'application/json'
        }

        if debug_metadata:
            print(f"[=] Testing accessibility of column '{internal_name}'...")

        response = make_graph_request_with_retry(url, headers, params=params)

        if response.status_code == 200:
            # Query succeeded - column is accessible
            data = response.json()
            items = data.get('value', [])

            if items and 'fields' in items[0]:
                # Check if column appears in the fields
                fields = items[0]['fields']
                column_in_fields = internal_name in fields or any(k.lower() == internal_name.lower() for k in fields.keys())

                if debug_metadata:
                    if column_in_fields:
                        print(f"[OK] Column '{internal_name}' found in item fields")
                    else:
                        print(f"[=] Column '{internal_name}' not in first item (may be new/empty)")

                # Return True regardless - if query succeeds, column is accessible
                # It just might not have values yet
                return True

            # No items in list - column still accessible (list is empty)
            if debug_metadata:
                print(f"[OK] Column '{internal_name}' accessible (list has no items yet)")
            return True
        else:
            if debug_metadata:
                print(f"[!] Column '{internal_name}' accessibility test failed: {response.status_code}")
            return False

    except Exception as e:
        if is_debug_metadata_enabled():
            print(f"[!] Error testing column accessibility: {e}")
        return False


def comprehensive_column_verification(site_id, list_id, token, graph_endpoint, column_name):
    """
    Comprehensive verification of column existence and properties.

    Performs detailed analysis of a SharePoint column including:
    - Existence verification
    - Property inspection (type, required, indexed, etc.)
    - Accessibility testing
    - Type-specific property analysis

    Args:
        site_id (str): SharePoint site ID
        list_id (str): SharePoint list/library ID
        token (str): OAuth access token
        graph_endpoint (str): Microsoft Graph API endpoint
        column_name (str): Name of column to verify (display or internal name)

    Returns:
        dict: Column analysis dictionary with properties, or None if not found
              Format: {
                  'exists': bool,
                  'display_name': str,
                  'internal_name': str,
                  'id': str,
                  'description': str,
                  'type': str,
                  'required': bool,
                  'hidden': bool,
                  'indexed': bool,
                  'read_only': bool,
                  'enforce_unique': bool,
                  'accessible': bool,
                  'text_properties': dict (if type is text)
              }

    Note:
        Results include detailed property inspection and accessibility testing.
        Use verify_column_for_filehash_operations() for FileHash-specific validation.
    """
    try:
        debug_metadata = is_debug_metadata_enabled()

        if debug_metadata:
            print(f"[=] Starting comprehensive verification for column '{column_name}'")

        # Step 1: Get all columns with detailed properties
        url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists/{list_id}/columns"
        headers = {
            'Authorization': f'Bearer {token}',
            'Accept': 'application/json'
        }

        response = make_graph_request_with_retry(url, headers)

        if response.status_code != 200:
            print(f"[!] Failed to retrieve columns: {response.status_code}")
            if debug_metadata:
                print(f"[DEBUG] Response: {response.text[:500]}")
            return None

        columns_data = response.json().get('value', [])
        target_column = None

        # Step 2: Find target column by name or display name
        for column in columns_data:
            if (column.get('name', '').lower() == column_name.lower() or
                column.get('displayName', '').lower() == column_name.lower()):
                target_column = column
                break

        if not target_column:
            print(f"[!] Column '{column_name}' not found in list")
            if debug_metadata:
                available_columns = [col.get('displayName', col.get('name', 'N/A')) for col in columns_data[:10]]
                print(f"[DEBUG] Available columns (first 10): {available_columns}")
            return None

        # Debug: Show raw column data
        if debug_metadata:
            print(f"[DEBUG] Raw column data from Graph API:")
            print(f"[DEBUG] Column keys: {list(target_column.keys())}")
            # Show which type property exists
            type_props = [k for k in target_column.keys() if k in ['text', 'number', 'dateTime', 'boolean', 'choice', 'lookup', 'calculated']]
            if type_props:
                print(f"[DEBUG] Type properties found: {type_props}")

        # Step 3: Analyze column properties
        # Determine column type by checking which type-specific property exists
        column_type = ''
        if 'text' in target_column:
            column_type = 'text'
        elif 'number' in target_column:
            column_type = 'number'
        elif 'dateTime' in target_column:
            column_type = 'dateTime'
        elif 'boolean' in target_column:
            column_type = 'boolean'
        elif 'choice' in target_column:
            column_type = 'choice'
        elif 'lookup' in target_column:
            column_type = 'lookup'
        elif 'calculated' in target_column:
            column_type = 'calculated'

        column_analysis = {
            'exists': True,
            'display_name': target_column.get('displayName', ''),
            'internal_name': target_column.get('name', ''),
            'id': target_column.get('id', ''),
            'description': target_column.get('description', ''),
            'type': column_type,  # Detected from which property exists
            'required': target_column.get('required', False),
            'hidden': target_column.get('hidden', False),
            'indexed': target_column.get('indexed', False),
            'read_only': target_column.get('readOnly', False),
            'enforce_unique': target_column.get('enforceUniqueValues', False)
        }

        # Step 4: Type-specific analysis
        if 'text' in target_column:
            text_props = target_column['text']
            column_analysis['text_properties'] = {
                'max_length': text_props.get('maxLength', 0),
                'allow_multiple_lines': text_props.get('allowMultipleLines', False),
                'append_changes': text_props.get('appendChangesToExistingText', False)
            }

        # Step 5: Validate column accessibility
        if debug_metadata:
            print(f"[=] Testing column accessibility...")

        accessibility_test = test_column_accessibility(
            site_id, list_id, token, graph_endpoint, column_analysis['internal_name']
        )
        column_analysis['accessible'] = accessibility_test

        # Step 6: Report findings
        if debug_metadata:
            print(f"\n" + "="*40)
            print(f"COLUMN VERIFICATION REPORT")
            print("="*40)
            print(f"Display Name: {column_analysis['display_name']}")
            print(f"Internal Name: {column_analysis['internal_name']}")
            print(f"Type: {column_analysis['type']}")
            print(f"Required: {column_analysis['required']}")
            print(f"Hidden: {column_analysis['hidden']}")
            print(f"Indexed: {column_analysis['indexed']}")
            print(f"Read Only: {column_analysis['read_only']}")
            print(f"Enforce Unique: {column_analysis['enforce_unique']}")

            if 'text_properties' in column_analysis:
                text_props = column_analysis.get('text_properties', {})
                print(f"Max Length: {text_props.get('max_length', 'N/A')}")
                print(f"Multiple Lines: {text_props.get('allow_multiple_lines', 'N/A')}")

            print(f"Accessible: {column_analysis['accessible']}")

            if column_analysis['hidden']:
                print(f"[⚠] WARNING: Column is hidden")
            if column_analysis['read_only']:
                print(f"[⚠] WARNING: Column is read-only")
            if not column_analysis['accessible']:
                print(f"[!] ERROR: Column exists but is not accessible")

            print("="*40 + "\n")

        return column_analysis

    except Exception as e:
        print(f"[!] Error in comprehensive column verification: {e}")
        if is_debug_metadata_enabled():
            import traceback
            print(f"[DEBUG] Full traceback: {traceback.format_exc()}")
        return None


def verify_column_for_filehash_operations(site_id, list_id, token, graph_endpoint):
    """
    Specific verification for FileHash column operations.

    Validates that the FileHash column exists, is accessible, and is suitable
    for storing xxHash128 checksums (32-character hex strings).

    Args:
        site_id (str): SharePoint site ID
        list_id (str): SharePoint list/library ID
        token (str): OAuth access token
        graph_endpoint (str): Microsoft Graph API endpoint

    Returns:
        tuple: (is_valid: bool, message: str)
               - is_valid: True if FileHash column is suitable for operations
               - message: Description of validation result or issues found

    Note:
        This function checks:
        - Column existence
        - Column accessibility
        - Column type (must be 'text')
        - Max length (must accommodate 32 characters)
        - Read-only status (must be writable)
        - Hidden status (should not be hidden)
    """
    try:
        debug_metadata = is_debug_metadata_enabled()

        if debug_metadata:
            print(f"[=] Verifying FileHash column for operations...")

        verification_result = comprehensive_column_verification(
            site_id, list_id, token, graph_endpoint, "FileHash"
        )

        if not verification_result:
            return False, "Column not found"

        # Check if suitable for hash storage
        issues = []

        if verification_result.get('read_only', False):
            issues.append("Column is read-only")

        if verification_result.get('hidden', False):
            issues.append("Column is hidden")

        if not verification_result.get('accessible', False):
            issues.append("Column is not accessible")

        if verification_result.get('type', '') != 'text':
            issues.append(f"Column type is {verification_result.get('type', 'unknown')}, expected 'text'")

        text_props = verification_result.get('text_properties', {})
        max_length = text_props.get('max_length', 0)
        if 0 < max_length < 32:
            issues.append(f"Max length ({max_length}) too small for hash (needs 32)")

        if issues:
            if debug_metadata:
                print(f"[!] FileHash column issues found:")
                for issue in issues:
                    print(f"    - {issue}")
            return False, "; ".join(issues)

        if debug_metadata:
            print(f"[OK] FileHash column is suitable for operations")

        return True, "Column verified successfully"

    except Exception as e:
        error_msg = f"Error during verification: {str(e)[:200]}"
        if is_debug_metadata_enabled():
            print(f"[!] {error_msg}")
            import traceback
            print(f"[DEBUG] Full traceback: {traceback.format_exc()}")
        return False, error_msg


def create_graph_client(tenant_id, client_id, client_secret, login_endpoint, graph_endpoint):
    """
    DEPRECATED: This function is no longer used as of v4.0.0.

    The action now uses direct Graph REST API calls instead of Office365-REST-Python-Client.
    Use get_drive_item_by_path() to get SharePoint resources directly.

    Args:
        tenant_id: Not used (deprecated)
        client_id: Not used (deprecated)
        client_secret: Not used (deprecated)
        login_endpoint: Not used (deprecated)
        graph_endpoint: Not used (deprecated)

    Raises:
        NotImplementedError: Always raises as this function is deprecated
    """
    # Parameters intentionally unused - function is deprecated
    _ = (tenant_id, client_id, client_secret, login_endpoint, graph_endpoint)
    raise NotImplementedError(
        "create_graph_client() is deprecated as of v4.0.0. "
        "Use direct Graph REST API functions like get_drive_item_by_path() instead. "
        "The Office365-REST-Python-Client library has been removed."
    )


def list_files_in_folder_recursive(drive, folder_path, site_url, tenant_id, client_id,
                                   client_secret, login_endpoint, graph_endpoint, current_path=""):
    """
    Recursively list all files in a SharePoint folder using direct Graph REST API.

    Args:
        drive: Office365 Drive object representing the folder
        folder_path (str): The original folder path being synced
        site_url (str): SharePoint site URL
        tenant_id (str): Azure AD tenant ID
        client_id (str): Azure AD application client ID
        client_secret (str): Azure AD application client secret
        login_endpoint (str): Azure AD login endpoint
        graph_endpoint (str): Microsoft Graph API endpoint
        current_path (str): Current relative path within the folder structure

    Returns:
        list: List of dictionaries containing file information:
            - name (str): File name
            - path (str): Relative path from root folder
            - id (str): SharePoint item ID
            - size (int): File size in bytes
            - drive_item: The DriveItem object for deletion (None for Graph API)

    Note:
        Uses direct Graph REST API calls instead of Office365 library property detection
        to reliably distinguish between files and folders.
    """
    files = []
    debug_enabled = is_debug_enabled()

    try:
        # Get authentication token
        from .auth import acquire_token
        token = acquire_token(tenant_id, client_id, client_secret, login_endpoint, graph_endpoint)
        if not token:
            raise Exception("Failed to acquire authentication token")

        # Get site and drive IDs if this is the first call
        if not current_path:
            # Parse site URL to get site ID
            # Format: https://tenant.sharepoint.com/sites/sitename
            import urllib.parse
            parsed = urllib.parse.urlparse(site_url)
            hostname = parsed.netloc
            site_path = parsed.path

            # Get site ID
            site_id_url = f"https://{graph_endpoint}/v1.0/sites/{hostname}:{site_path}"
            headers = {
                'Authorization': f"Bearer {token['access_token']}",
                'Accept': 'application/json'
            }
            site_response = make_graph_request_with_retry(site_id_url, headers, method='GET')

            if site_response.status_code != 200:
                raise Exception(f"Failed to get site ID: {site_response.status_code} - {site_response.text}")

            site_data = site_response.json()
            site_id = site_data['id']

            if debug_enabled:
                print(f"[DEBUG] Site ID: {site_id}")

            # Get default drive ID
            drive_url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/drive"
            drive_response = make_graph_request_with_retry(drive_url, headers, method='GET')

            if drive_response.status_code != 200:
                raise Exception(f"Failed to get drive: {drive_response.status_code} - {drive_response.text}")

            drive_data = drive_response.json()
            drive_id = drive_data['id']

            if debug_enabled:
                print(f"[DEBUG] Drive ID: {drive_id}")

            # Get the folder item by path
            # URL encode the folder path
            encoded_path = urllib.parse.quote(folder_path.strip('/'))
            folder_url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}"
            folder_response = make_graph_request_with_retry(folder_url, headers, method='GET')

            if folder_response.status_code != 200:
                raise Exception(f"Failed to get folder: {folder_response.status_code} - {folder_response.text}")

            folder_data = folder_response.json()
            folder_item_id = folder_data['id']

            if debug_enabled:
                print(f"[DEBUG] Folder item ID: {folder_item_id}")

            # Store these for recursive calls in global cache
            site_drive_id_cache['site_id'] = site_id
            site_drive_id_cache['drive_id'] = drive_id
            site_drive_id_cache['current_item_id'] = folder_item_id
        else:
            # Use stored IDs from parent call
            site_id = site_drive_id_cache.get('site_id')
            drive_id = site_drive_id_cache.get('drive_id')
            folder_item_id = site_drive_id_cache.get('current_item_id')

        # Get children of the current folder using Graph API
        children_url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/drives/{drive_id}/items/{folder_item_id}/children"
        headers = {
            'Authorization': f"Bearer {token['access_token']}",
            'Accept': 'application/json'
        }

        children_response = make_graph_request_with_retry(children_url, headers, method='GET')

        if children_response.status_code != 200:
            raise Exception(f"Failed to list children: {children_response.status_code} - {children_response.text}")

        children_data = children_response.json()
        children = children_data.get('value', [])

        if debug_enabled and not current_path:
            print(f"\n[DEBUG] SharePoint folder contains {len(children)} items")

        for child in children:
            # Build the relative path for this item
            item_name = child.get('name', '')
            item_path = f"{current_path}/{item_name}" if current_path else item_name

            # Check if this item has a 'file' or 'folder' facet in the JSON
            has_file = 'file' in child
            has_folder = 'folder' in child

            if debug_enabled:
                item_type = "FILE" if has_file else ("FOLDER" if has_folder else "UNKNOWN")
                print(f"[DEBUG] SharePoint item: {item_path} (type: {item_type})")

            # If it's a file, add to list
            if has_file:
                file_info = {
                    'name': item_name,
                    'path': item_path,
                    'id': child.get('id', ''),
                    'size': child.get('size', 0),
                    'drive_item': None  # Graph API doesn't use Office365 drive_item objects
                }
                files.append(file_info)

                if debug_enabled:
                    print(f"  [+] Added to file list: {item_path} ({file_info['size']} bytes)")

            # If it's a folder, recurse into it
            elif has_folder:
                if debug_enabled:
                    print(f"  [→] Entering subfolder: {item_path}")

                # Store the child item ID in the cache for the recursive call
                child_item_id = child.get('id', '')
                previous_item_id = site_drive_id_cache.get('current_item_id')
                site_drive_id_cache['current_item_id'] = child_item_id

                # Recursively get files from this subfolder
                subfolder_files = list_files_in_folder_recursive(
                    drive, folder_path, site_url, tenant_id, client_id,
                    client_secret, login_endpoint, graph_endpoint, item_path
                )
                files.extend(subfolder_files)

                # Restore parent folder's item ID after recursion
                site_drive_id_cache['current_item_id'] = previous_item_id

                if debug_enabled:
                    print(f"  [←] Exited subfolder: {item_path} (found {len(subfolder_files)} files)")
            else:
                if debug_enabled:
                    print(f"  [!] WARNING: Item is neither file nor folder: {item_path}")

    except Exception as e:
        print(f"[!] Error listing files in folder '{current_path}': {str(e)}")
        if is_debug_metadata_enabled():
            import traceback
            print(f"[DEBUG] Traceback: {traceback.format_exc()}")

    # Debug summary for root folder only
    if debug_enabled and not current_path and len(files) > 0:
        print(f"[DEBUG] Returning {len(files)} FILES (folders excluded)")
        print(f"[DEBUG] Sample files (first 5):")
        for f in files[:5]:
            print(f"  - {f['path']} ({f['size']} bytes)")

    return files


def build_sharepoint_cache(folder_path, site_url, tenant_id, client_id,
                          client_secret, login_endpoint, graph_endpoint,
                          filehash_available=True, current_path="", parent_item_id=None,
                          folder_cache_dict=None):
    """
    Build a comprehensive cache of all files and folders in SharePoint with metadata.

    This function performs a single recursive walk of the SharePoint folder structure
    and retrieves all file metadata including FileHash values and folder item IDs in
    one operation. This dramatically reduces API calls compared to querying each
    file/folder individually.

    Performance Benefits:
        - 80-90% reduction in API calls for file comparisons
        - 90-95% reduction in folder existence check API calls
        - 15-20 seconds faster for typical 100-file repository
        - Eliminates per-file network latency
        - Reusable for both file comparison and sync deletion

    Args:
        folder_path (str): The SharePoint folder path to cache (e.g., "Documents/Folder")
        site_url (str): SharePoint site URL (e.g., "https://company.sharepoint.com/sites/site")
        tenant_id (str): Azure AD tenant ID
        client_id (str): Azure AD application client ID
        client_secret (str): Azure AD application client secret
        login_endpoint (str): Azure AD login endpoint
        graph_endpoint (str): Microsoft Graph API endpoint
        filehash_available (bool): Whether FileHash column exists (default: True)
        current_path (str): Internal - current relative path during recursion
        parent_item_id (str): Internal - parent folder item ID during recursion
        folder_cache_dict (dict): Internal - shared folder cache across recursive calls

    Returns:
        dict: Dictionary with 'files' and 'folders' keys:
            {
                'files': {
                    "path/to/file.html": {
                        "item_id": "abc123",           # Drive item ID
                        "list_item_id": "def456",      # List item ID (for metadata updates)
                        "parent_item_id": "xyz789",    # Parent folder item ID
                        "file_hash": "a1b2c3d4...",   # FileHash column value (if available)
                        "size": 12345,                 # File size in bytes
                        "name": "file.html"            # File name
                    },
                    ...
                },
                'folders': {
                    "path/to/folder": {
                        "item_id": "xyz789",           # Folder drive item ID
                        "name": "folder"               # Folder name
                    },
                    ...
                }
            }

    Graph API Query:
        Uses: /children?$expand=listItem($expand=fields($select=FileHash,FileSizeDisplay,FileLeafRef))

        This retrieves in a single call:
        - Drive item metadata (id, name, size)
        - List item ID (for metadata updates)
        - Custom column values (FileHash)
        - Standard fields (FileSizeDisplay)

    Cache Miss Handling:
        Functions using the cache should fall back to individual API queries if:
        - File not found in cache (newly created during execution)
        - Cache is None (caching disabled for force upload)

    Note:
        - Cache is built once at beginning of execution
        - Cache does NOT auto-update during execution
        - Cache may become stale if files are modified during execution (rare in CI/CD)

    """
    cache = {}

    # Initialize folder cache on first call (root level)
    if folder_cache_dict is None:
        folder_cache_dict = {}

    debug_enabled = is_debug_enabled()
    debug_metadata = is_debug_metadata_enabled()

    try:
        # Get authentication token
        from .auth import acquire_token
        token = acquire_token(tenant_id, client_id, client_secret, login_endpoint, graph_endpoint)
        if not token:
            raise Exception("Failed to acquire authentication token for cache building")

        # First call: Initialize site/drive IDs and get root folder item ID
        if not current_path:
            print(f"[*] Building SharePoint metadata cache for: {folder_path}")

            # Parse site URL to get site ID
            import urllib.parse
            parsed = urllib.parse.urlparse(site_url)
            hostname = parsed.netloc
            site_path = parsed.path

            # Get site ID
            site_id_url = f"https://{graph_endpoint}/v1.0/sites/{hostname}:{site_path}"
            headers = {
                'Authorization': f"Bearer {token['access_token']}",
                'Accept': 'application/json'
            }
            site_response = make_graph_request_with_retry(site_id_url, headers, method='GET')

            if site_response.status_code != 200:
                raise Exception(f"Failed to get site ID: {site_response.status_code}")

            site_data = site_response.json()
            site_id = site_data['id']

            # Get default drive ID
            drive_url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/drive"
            drive_response = make_graph_request_with_retry(drive_url, headers, method='GET')

            if drive_response.status_code != 200:
                raise Exception(f"Failed to get drive: {drive_response.status_code}")

            drive_data = drive_response.json()
            drive_id = drive_data['id']

            # Get the folder item by path
            encoded_path = urllib.parse.quote(folder_path.strip('/'))
            folder_url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}"
            folder_response = make_graph_request_with_retry(folder_url, headers, method='GET')

            if folder_response.status_code != 200:
                raise Exception(f"Failed to get folder: {folder_response.status_code}")

            folder_data = folder_response.json()
            folder_item_id = folder_data['id']

            # Store in cache for recursive calls
            site_drive_id_cache['site_id'] = site_id
            site_drive_id_cache['drive_id'] = drive_id
            site_drive_id_cache['current_item_id'] = folder_item_id

            parent_item_id = folder_item_id

            if debug_metadata:
                print(f"[DEBUG] Cache builder - Site ID: {site_id}")
                print(f"[DEBUG] Cache builder - Drive ID: {drive_id}")
                print(f"[DEBUG] Cache builder - Root folder item ID: {folder_item_id}")
        else:
            # Recursive call: Use cached IDs
            site_id = site_drive_id_cache.get('site_id')
            drive_id = site_drive_id_cache.get('drive_id')
            folder_item_id = parent_item_id or site_drive_id_cache.get('current_item_id')

        # Build children URL with listItem expansion to get metadata in one call
        # Syntax: $expand=listItem($expand=fields($select=Field1,Field2))
        # Note: Semicolon (;) separates $select and $expand within nested parameters
        if filehash_available:
            # Include FileHash in field selection
            expand_clause = "listItem($expand=fields($select=FileHash,FileSizeDisplay,FileLeafRef))"
        else:
            # Skip FileHash if column doesn't exist
            expand_clause = "listItem($expand=fields($select=FileSizeDisplay,FileLeafRef))"

        children_url = (f"https://{graph_endpoint}/v1.0/sites/{site_id}/drives/{drive_id}"
                       f"/items/{folder_item_id}/children?$expand={expand_clause}")

        headers = {
            'Authorization': f"Bearer {token['access_token']}",
            'Accept': 'application/json'
        }

        if debug_metadata:
            print(f"[DEBUG] Cache query: {children_url}")

        children_response = make_graph_request_with_retry(children_url, headers, method='GET')

        if children_response.status_code != 200:
            print(f"[!] Warning: Failed to list children for cache: {children_response.status_code}")
            if debug_metadata:
                print(f"[DEBUG] Response: {children_response.text[:500]}")
            return cache  # Return empty cache on error

        children_data = children_response.json()
        children = children_data.get('value', [])

        if debug_enabled and not current_path:
            print(f"[*] Found {len(children)} items in root folder")

        # Process each child item
        for child in children:
            item_name = child.get('name', '')
            item_path = f"{current_path}/{item_name}" if current_path else item_name

            # Check if this is a file or folder
            has_file = 'file' in child
            has_folder = 'folder' in child

            if has_file:
                # Extract metadata from the response
                item_id = child.get('id', '')
                size = child.get('size', 0)

                # Extract list item data if available
                list_item = child.get('listItem')
                list_item_id = None
                file_hash = None

                if list_item:
                    list_item_id = list_item.get('id')
                    fields = list_item.get('fields', {})

                    if fields:
                        # Get FileHash if column is available
                        if filehash_available:
                            file_hash = fields.get('FileHash')

                        if debug_metadata and not current_path and len(cache) < 3:
                            # Show first few files as examples
                            print(f"[DEBUG] Cached file: {item_path}")
                            print(f"[DEBUG]   - item_id: {item_id}")
                            print(f"[DEBUG]   - list_item_id: {list_item_id}")
                            print(f"[DEBUG]   - size: {size}")
                            if file_hash:
                                print(f"[DEBUG]   - file_hash: {file_hash[:16]}...")

                # Add to cache
                cache[item_path] = {
                    'item_id': item_id,
                    'list_item_id': list_item_id,
                    'parent_item_id': folder_item_id,
                    'file_hash': file_hash,
                    'size': size,
                    'name': item_name
                }

            elif has_folder:
                # Recurse into subfolder
                if debug_enabled:
                    print(f"[*] Caching subfolder: {item_path}")

                child_item_id = child.get('id', '')

                # Add folder to folder cache
                folder_cache_dict[item_path] = {
                    'item_id': child_item_id,
                    'name': item_name
                }

                # Recursive call with child folder's item ID
                # Pass folder_cache_dict so all folders accumulate in same dictionary
                subfolder_cache = build_sharepoint_cache(
                    folder_path, site_url, tenant_id, client_id,
                    client_secret, login_endpoint, graph_endpoint,
                    filehash_available, item_path, child_item_id,
                    folder_cache_dict
                )

                # Merge subfolder cache into main cache
                cache.update(subfolder_cache)

    except Exception as e:
        print(f"[!] Error building SharePoint cache for '{current_path}': {str(e)}")
        if debug_metadata:
            import traceback
            print(f"[DEBUG] Traceback: {traceback.format_exc()}")

    # Summary for root folder only (always show, not just in debug mode)
    if not current_path and len(cache) > 0:
        print()
        print("[CACHE] SharePoint Metadata Cache:")
        print(f"   - Total files cached:       {len(cache):>6}")
        print(f"   - Total folders cached:     {len(folder_cache_dict):>6}")

        # Show statistics
        files_with_hash = sum(1 for f in cache.values() if f.get('file_hash'))
        files_with_list_id = sum(1 for f in cache.values() if f.get('list_item_id'))

        if filehash_available:
            print(f"   - Files with FileHash:      {files_with_hash:>6}/{len(cache)}")
        print(f"   - Files with list_item_id:  {files_with_list_id:>6}/{len(cache)}")

        if debug_metadata and len(cache) > 0:
            print(f"[DEBUG] Sample cached file paths (first 5):")
            for path in list(cache.keys())[:5]:
                print(f"  - {path}")
            if len(folder_cache_dict) > 0:
                print(f"[DEBUG] Sample cached folder paths (first 5):")
                for path in list(folder_cache_dict.keys())[:5]:
                    print(f"  - {path}")

    # Return combined cache with files and folders
    # On root call, return the structured dictionary
    if not current_path:
        return {
            'files': cache,
            'folders': folder_cache_dict
        }
    else:
        # On recursive calls, return just the file cache
        # (folder cache is accumulated through folder_cache_dict parameter)
        return cache


def delete_file_from_sharepoint(drive_item, file_path, whatif=False, file_id=None,
                               site_url=None, tenant_id=None, client_id=None,
                               client_secret=None, login_endpoint=None, graph_endpoint=None):
    """
    Delete a file from SharePoint.

    Args:
        drive_item: Office365 DriveItem object representing the file to delete (or None for Graph API)
        file_path (str): Relative path of the file (for logging)
        whatif (bool): If True, simulate deletion without actually deleting (default: False)
        file_id (str): SharePoint item ID for Graph API deletion (required if drive_item is None)
        site_url (str): SharePoint site URL (required if drive_item is None)
        tenant_id (str): Azure AD tenant ID (required if drive_item is None)
        client_id (str): Azure AD application client ID (required if drive_item is None)
        client_secret (str): Azure AD application client secret (required if drive_item is None)
        login_endpoint (str): Azure AD login endpoint (required if drive_item is None)
        graph_endpoint (str): Microsoft Graph API endpoint (required if drive_item is None)

    Returns:
        bool: True if deletion successful (or would be successful in whatif mode), False otherwise

    Note:
        Supports two modes:
        1. Office365 library (drive_item provided) - legacy mode
        2. Direct Graph API (drive_item=None, file_id provided) - new mode
        WhatIf mode allows users to preview deletions without actually performing them.

        site_url parameter is not used in Graph API mode (uses cached site_id/drive_id instead)
        but kept in signature for API compatibility.
    """
    # site_url intentionally unused - we use cached site_id/drive_id from global cache
    _ = site_url

    debug_enabled = is_debug_enabled()

    try:
        if whatif:
            # WhatIf mode - just show what would be deleted
            print(f"File Deleted (WhatIf): {file_path}")
            if debug_enabled:
                print(f"  → Would delete this file (WhatIf mode active)")
            return True
        else:
            # Actually delete the file
            if debug_enabled:
                print(f"[×] Deleting file from SharePoint: {file_path}")

            # Check which deletion method to use
            if drive_item is None:
                # Use Graph API deletion
                if not file_id:
                    raise Exception("file_id is required for Graph API deletion")

                # Get authentication token
                from .auth import acquire_token
                token = acquire_token(tenant_id, client_id, client_secret, login_endpoint, graph_endpoint)
                if not token:
                    raise Exception("Failed to acquire authentication token")

                # Use stored site and drive IDs from global cache (set by list_files_in_folder_recursive)
                site_id = site_drive_id_cache.get('site_id')
                drive_id = site_drive_id_cache.get('drive_id')

                # Delete the file using Graph API
                delete_url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/drives/{drive_id}/items/{file_id}"
                headers = {
                    'Authorization': f"Bearer {token['access_token']}",
                    'Accept': 'application/json'
                }

                delete_response = make_graph_request_with_retry(delete_url, headers, method='DELETE')

                if delete_response.status_code not in [200, 204]:
                    raise Exception(f"Failed to delete file: {delete_response.status_code} - {delete_response.text}")

            else:
                # Use Office365 library deletion (legacy mode)
                drive_item.delete_object().execute_query()

            # Always show simple deletion message
            print(f"File Deleted: {file_path}")

            # Show detailed message only in DEBUG mode
            if debug_enabled:
                print(f"  → Deletion confirmed")

            return True

    except Exception as e:
        print(f"[!] Failed to delete file '{file_path}': {str(e)}")
        if is_debug_metadata_enabled():
            import traceback
            print(f"[DEBUG] Traceback: {traceback.format_exc()}")
        return False


def get_drive_item_by_path(site_url, folder_path, tenant_id, client_id,
                           client_secret, login_endpoint, graph_endpoint, 
                           specify_upload_library):
    """
    Get a drive item (file or folder) by its path using Graph API.

    Args:
        site_url (str): SharePoint site URL
        folder_path (str): Path to the item (e.g., 'Documents/Folder1/file.txt')
        tenant_id (str): Azure AD tenant ID
        client_id (str): Azure AD application client ID
        client_secret (str): Azure AD application client secret
        login_endpoint (str): Azure AD login endpoint
        graph_endpoint (str): Microsoft Graph API endpoint

    Returns:
        dict: Drive item metadata including:
            - id: Item ID
            - name: Item name
            - size: Item size
            - webUrl: Web URL
            - (and other driveItem properties)
        None: If item not found or error occurred

    Example:
        item = get_drive_item_by_path(site_url, 'Documents/Reports', ...)
        item_id = item['id']
    """
    debug_enabled = is_debug_enabled()

    try:
        # Get authentication token
        from .auth import acquire_token
        token = acquire_token(tenant_id, client_id, client_secret, login_endpoint, graph_endpoint)
        if not token:
            raise Exception("Failed to acquire authentication token")

        # Parse site URL to get site ID
        import urllib.parse
        parsed = urllib.parse.urlparse(site_url)
        hostname = parsed.netloc
        site_path = parsed.path

        # Get site ID
        site_id_url = f"https://{graph_endpoint}/v1.0/sites/{hostname}:{site_path}"
        headers = {
            'Authorization': f"Bearer {token['access_token']}",
            'Accept': 'application/json'
        }
        site_response = make_graph_request_with_retry(site_id_url, headers, method='GET')

        if site_response.status_code != 200:
            raise Exception(f"Failed to get site ID: {site_response.status_code}")

        site_data = site_response.json()
        site_id = site_data['id']

        if specify_upload_library:
            library_name = folder_path.split('/')[0]
            drive_url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/drives"
            drives_response = make_graph_request_with_retry(drive_url, headers, method='GET')

            if drives_response.status_code != 200:
                raise Exception(f"Failed to get drives: {drives_response.status_code}")
            
            drives_data = drives_response.json()
            drives = drives_data.get('value', [])
            drive_id = next((drive['id'] for drive in drives if drive.get('name') == library_name), None)


        else:
            # Get default drive ID
            drive_url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/drive"
            drive_response = make_graph_request_with_retry(drive_url, headers, method='GET')

            if drive_response.status_code != 200:
                raise Exception(f"Failed to get drive: {drive_response.status_code}")

            drive_data = drive_response.json()
            drive_id = drive_data['id']

        # Get the item by path
        if specify_upload_library:
            encoded_path = urllib.parse.quote('/'.join(folder_path.split('/')[1:]).strip('/'))
        else:
            encoded_path = urllib.parse.quote(folder_path.strip('/'))
        item_url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}"

        item_response = make_graph_request_with_retry(item_url, headers, method='GET')

        if item_response.status_code == 200:
            item_data = item_response.json()
            # Store IDs for other functions to use
            item_data['_site_id'] = site_id
            item_data['_drive_id'] = drive_id
            return item_data
        elif item_response.status_code == 404:
            if debug_enabled:
                print(f"[!] Item not found: {folder_path}")
            return None
        else:
            raise Exception(f"Failed to get item: {item_response.status_code} - {item_response.text}")

    except Exception as e:
        print(f"[!] Error getting drive item by path: {str(e)}")
        if is_debug_metadata_enabled():
            import traceback
            print(f"[DEBUG] Traceback: {traceback.format_exc()}")
        return None


def get_drive_item_by_path_with_list_item(site_id, drive_id, parent_item_id, filename,
                                           tenant_id, client_id, client_secret, login_endpoint, graph_endpoint):
    """
    Get a drive item by path with its list item property expanded.

    This is the most direct way to get the list item ID after upload, since we know
    exactly where we uploaded the file (parent folder + filename).

    Args:
        site_id (str): SharePoint site ID
        drive_id (str): SharePoint drive ID
        parent_item_id (str): Parent folder item ID
        filename (str): Filename (should be URL-encoded if contains special chars)
        tenant_id (str): Azure AD tenant ID
        client_id (str): Azure AD application client ID
        client_secret (str): Azure AD application client secret
        login_endpoint (str): Azure AD login endpoint
        graph_endpoint (str): Microsoft Graph API endpoint

    Returns:
        dict: Drive item with listItem property containing the list item ID
        None: If fetch failed

    Example:
        After uploading to /items/ABC123:/{filename}:/content
        We can get it at /items/ABC123:/{filename}?$expand=listItem
    """
    try:
        # Get authentication token
        from .auth import acquire_token
        token = acquire_token(tenant_id, client_id, client_secret, login_endpoint, graph_endpoint)
        if not token:
            return None

        # URL encode the filename (should already be encoded but ensure it)
        import urllib.parse
        encoded_filename = urllib.parse.quote(filename)

        # Fetch drive item by path with listItem expanded
        # Uses the same path structure as upload: /items/{parent-id}:/{filename}
        item_url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/drives/{drive_id}/items/{parent_item_id}:/{encoded_filename}?$expand=listItem"

        headers = {
            'Authorization': f"Bearer {token['access_token']}",
            'Accept': 'application/json'
        }

        if is_debug_enabled():
            print(f"[DEBUG] Fetching drive item by path with listItem")

        response = make_graph_request_with_retry(item_url, headers, method='GET')

        if response.status_code == 200:
            return response.json()
        else:
            if is_debug_enabled():
                print(f"[DEBUG] Failed to fetch drive item by path: {response.status_code} - {response.text[:200]}")
            return None

    except Exception as e:
        if is_debug_enabled():
            print(f"[DEBUG] Error fetching drive item by path: {str(e)[:200]}")
        return None


def get_drive_item_with_list_item(site_id, drive_id, item_id,
                                   tenant_id, client_id, client_secret, login_endpoint, graph_endpoint):
    """
    Get a drive item by ID with its list item property expanded.

    This is a fallback method when we have the drive item ID but not the path.
    Prefer using get_drive_item_by_path_with_list_item() when possible since
    we usually know the exact path where we uploaded.

    Args:
        site_id (str): SharePoint site ID
        drive_id (str): SharePoint drive ID
        item_id (str): Drive item ID (file ID in the drive)
        tenant_id (str): Azure AD tenant ID
        client_id (str): Azure AD application client ID
        client_secret (str): Azure AD application client secret
        login_endpoint (str): Azure AD login endpoint
        graph_endpoint (str): Microsoft Graph API endpoint

    Returns:
        dict: Drive item with listItem property
        None: If fetch failed
    """
    try:
        # Get authentication token
        from .auth import acquire_token
        token = acquire_token(tenant_id, client_id, client_secret, login_endpoint, graph_endpoint)
        if not token:
            return None

        # Fetch drive item with listItem expanded
        item_url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/drives/{drive_id}/items/{item_id}?$expand=listItem"

        headers = {
            'Authorization': f"Bearer {token['access_token']}",
            'Accept': 'application/json'
        }

        response = make_graph_request_with_retry(item_url, headers, method='GET')

        if response.status_code == 200:
            return response.json()
        else:
            if is_debug_enabled():
                print(f"[DEBUG] Failed to fetch drive item by ID: {response.status_code} - {response.text[:200]}")
            return None

    except Exception as e:
        if is_debug_enabled():
            print(f"[DEBUG] Error fetching drive item by ID: {str(e)[:200]}")
        return None


def upload_small_file_graph(site_id, drive_id, parent_item_id, filename, file_content,
                            tenant_id, client_id, client_secret, login_endpoint, graph_endpoint):
    """
    Upload a small file (<250 MB) using Graph API.

    Args:
        site_id (str): SharePoint site ID
        drive_id (str): SharePoint drive ID
        parent_item_id (str): Parent folder item ID
        filename (str): Name for the uploaded file
        file_content (bytes): File content as bytes
        tenant_id (str): Azure AD tenant ID
        client_id (str): Azure AD application client ID
        client_secret (str): Azure AD application client secret
        login_endpoint (str): Azure AD login endpoint
        graph_endpoint (str): Microsoft Graph API endpoint

    Returns:
        dict: Uploaded drive item metadata
        None: If upload failed

    Note:
        This method only supports files up to 250 MB in size.
        For larger files, use create_upload_session_graph().
    """
    debug_enabled = is_debug_enabled()

    try:
        # Get authentication token
        from .auth import acquire_token
        token = acquire_token(tenant_id, client_id, client_secret, login_endpoint, graph_endpoint)
        if not token:
            raise Exception("Failed to acquire authentication token")

        # URL encode the filename
        import urllib.parse
        encoded_filename = urllib.parse.quote(filename)

        # Upload endpoint: PUT /items/{parent-id}:/{filename}:/content
        # Note: Upload endpoint does not officially support $expand parameter
        # (Graph API error: "The type 'Edm.Stream' is not valid for $select or $expand")
        # We'll fetch the listItem separately after upload if needed
        upload_url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/drives/{drive_id}/items/{parent_item_id}:/{encoded_filename}:/content"

        headers = {
            'Authorization': f"Bearer {token['access_token']}",
            'Content-Type': 'application/octet-stream'
        }

        if debug_enabled:
            print(f"[DEBUG] Uploading to: {upload_url}")
            print(f"[DEBUG] File size: {len(file_content)} bytes")

        # Make the upload request (use data parameter for binary content)
        upload_response = make_graph_request_with_retry(upload_url, headers, method='PUT', data=file_content)

        if upload_response.status_code in [200, 201]:
            item_data = upload_response.json()
            if debug_enabled:
                print(f"[DEBUG] Upload successful: {item_data.get('id')}")
            return item_data
        else:
            raise Exception(f"Upload failed: {upload_response.status_code} - {upload_response.text}")

    except Exception as e:
        print(f"[!] Error uploading small file: {str(e)}")
        if is_debug_metadata_enabled():
            import traceback
            print(f"[DEBUG] Traceback: {traceback.format_exc()}")
        return None


def create_upload_session_graph(site_id, drive_id, parent_item_id, filename,
                                tenant_id, client_id, client_secret, login_endpoint, graph_endpoint):
    """
    Create an upload session for large files (>250 MB) using Graph API.

    Args:
        site_id (str): SharePoint site ID
        drive_id (str): SharePoint drive ID
        parent_item_id (str): Parent folder item ID
        filename (str): Name for the uploaded file
        tenant_id (str): Azure AD tenant ID
        client_id (str): Azure AD application client ID
        client_secret (str): Azure AD application client secret
        login_endpoint (str): Azure AD login endpoint
        graph_endpoint (str): Microsoft Graph API endpoint

    Returns:
        dict: Upload session info including:
            - uploadUrl: URL for uploading chunks
            - expirationDateTime: Session expiration
        None: If session creation failed

    Note:
        Use upload_file_chunk_graph() to upload chunks to the returned uploadUrl.
    """
    debug_enabled = is_debug_enabled()

    try:
        # Get authentication token
        from .auth import acquire_token
        token = acquire_token(tenant_id, client_id, client_secret, login_endpoint, graph_endpoint)
        if not token:
            raise Exception("Failed to acquire authentication token")

        # URL encode the filename
        import urllib.parse
        encoded_filename = urllib.parse.quote(filename)

        # Create upload session endpoint
        session_url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/drives/{drive_id}/items/{parent_item_id}:/{encoded_filename}:/createUploadSession"

        headers = {
            'Authorization': f"Bearer {token['access_token']}",
            'Content-Type': 'application/json'
        }

        # Request body with conflict behavior
        request_body = {
            "item": {
                "@microsoft.graph.conflictBehavior": "replace"
            }
        }

        if debug_enabled:
            print(f"[DEBUG] Creating upload session: {session_url}")

        session_response = make_graph_request_with_retry(session_url, headers, method='POST', json_data=request_body)

        if session_response.status_code == 200:
            session_data = session_response.json()
            if debug_enabled:
                print(f"[DEBUG] Upload session created: {session_data.get('uploadUrl')[:50]}...")
            return session_data
        else:
            raise Exception(f"Session creation failed: {session_response.status_code} - {session_response.text}")

    except Exception as e:
        print(f"[!] Error creating upload session: {str(e)}")
        if is_debug_metadata_enabled():
            import traceback
            print(f"[DEBUG] Traceback: {traceback.format_exc()}")
        return None


def upload_file_chunk_graph(upload_url, chunk_data, chunk_start, chunk_end, total_size):
    """
    Upload a chunk of a file to an upload session using Graph API.

    Args:
        upload_url (str): Upload URL from create_upload_session_graph()
        chunk_data (bytes): Chunk content as bytes
        chunk_start (int): Starting byte position (0-indexed)
        chunk_end (int): Ending byte position (inclusive)
        total_size (int): Total file size in bytes

    Returns:
        dict: Upload response:
            - If chunk accepted: Status info with nextExpectedRanges
            - If upload complete: Full driveItem metadata
        None: If upload failed

    Note:
        - Chunk sizes must be multiples of 320 KiB (327,680 bytes)
        - Maximum 60 MiB per chunk
        - Content-Range format: "bytes {start}-{end}/{total}"
    """
    debug_enabled = is_debug_enabled()

    try:
        import requests

        headers = {
            'Content-Length': str(len(chunk_data)),
            'Content-Range': f"bytes {chunk_start}-{chunk_end}/{total_size}"
        }

        if debug_enabled:
            print(f"[DEBUG] Uploading chunk: bytes {chunk_start}-{chunk_end}/{total_size}")

        # Use requests directly (no retry for chunks per MS documentation)
        response = requests.put(upload_url, headers=headers, data=chunk_data, timeout=300)

        # Check response
        if response.status_code in [200, 201, 202]:
            # 202 = chunk accepted, more chunks expected
            # 200/201 = upload complete
            response_data = response.json() if response.content else {}
            if debug_enabled:
                if response.status_code == 202:
                    print(f"[DEBUG] Chunk accepted, continuing...")
                else:
                    print(f"[DEBUG] Upload complete!")
            return response_data
        else:
            raise Exception(f"Chunk upload failed: {response.status_code} - {response.text}")

    except Exception as e:
        print(f"[!] Error uploading chunk: {str(e)}")
        if is_debug_metadata_enabled():
            import traceback
            print(f"[DEBUG] Traceback: {traceback.format_exc()}")
        return None


def create_folder_graph(site_id, drive_id, parent_item_id, folder_name,
                       tenant_id, client_id, client_secret, login_endpoint, graph_endpoint):
    """
    Create a folder in SharePoint using Graph API.

    Args:
        site_id (str): SharePoint site ID
        drive_id (str): SharePoint drive ID
        parent_item_id (str): Parent folder item ID where new folder will be created
        folder_name (str): Name for the new folder
        tenant_id (str): Azure AD tenant ID
        client_id (str): Azure AD application client ID
        client_secret (str): Azure AD application client secret
        login_endpoint (str): Azure AD login endpoint
        graph_endpoint (str): Microsoft Graph API endpoint

    Returns:
        dict: Created folder driveItem metadata including:
            - id: Folder item ID
            - name: Folder name
            - folder: Folder properties
        None: If folder creation failed

    Note:
        If a folder with the same name exists, this will automatically
        rename the new folder (e.g., "Folder" -> "Folder 1").
    """
    debug_enabled = is_debug_enabled()

    try:
        # Get authentication token
        from .auth import acquire_token
        token = acquire_token(tenant_id, client_id, client_secret, login_endpoint, graph_endpoint)
        if not token:
            raise Exception("Failed to acquire authentication token")

        # Create folder endpoint: POST /items/{parent-id}/children
        create_url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/drives/{drive_id}/items/{parent_item_id}/children"

        headers = {
            'Authorization': f"Bearer {token['access_token']}",
            'Content-Type': 'application/json'
        }

        # Request body
        request_body = {
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename"
        }

        if debug_enabled:
            print(f"[DEBUG] Creating folder: {folder_name} in parent {parent_item_id}")

        create_response = make_graph_request_with_retry(create_url, headers, method='POST', json_data=request_body)

        if create_response.status_code in [200, 201]:
            folder_data = create_response.json()
            if debug_enabled:
                print(f"[DEBUG] Folder created: {folder_data.get('id')}")
            return folder_data
        else:
            raise Exception(f"Folder creation failed: {create_response.status_code} - {create_response.text}")

    except Exception as e:
        print(f"[!] Error creating folder: {str(e)}")
        if is_debug_metadata_enabled():
            import traceback
            print(f"[DEBUG] Traceback: {traceback.format_exc()}")
        return None


def list_folder_children_graph(site_id, drive_id, item_id, tenant_id, client_id,
                               client_secret, login_endpoint, graph_endpoint, folder_path=None):
    """
    List all children (files and folders) in a folder using Graph API.

    Args:
        site_id (str): SharePoint site ID
        drive_id (str): SharePoint drive ID
        item_id (str): Folder item ID to list children of
        tenant_id (str): Azure AD tenant ID
        client_id (str): Azure AD application client ID
        client_secret (str): Azure AD application client secret
        login_endpoint (str): Azure AD login endpoint
        graph_endpoint (str): Microsoft Graph API endpoint
        folder_path (str, optional): Human-readable folder path for debug output

    Returns:
        list: List of driveItem dictionaries, each with:
            - id: Item ID
            - name: Item name
            - file: File facet (if file)
            - folder: Folder facet (if folder)
        None: If listing failed

    Note:
        Use 'file' in item or 'folder' in item to determine type.
    """
    debug_enabled = is_debug_enabled()

    try:
        # Get authentication token
        from .auth import acquire_token
        token = acquire_token(tenant_id, client_id, client_secret, login_endpoint, graph_endpoint)
        if not token:
            raise Exception("Failed to acquire authentication token")

        # List children endpoint
        children_url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/drives/{drive_id}/items/{item_id}/children"

        headers = {
            'Authorization': f"Bearer {token['access_token']}",
            'Accept': 'application/json'
        }

        children_response = make_graph_request_with_retry(children_url, headers, method='GET')

        if children_response.status_code == 200:
            children_data = children_response.json()
            children = children_data.get('value', [])
            if debug_enabled:
                if folder_path:
                    print(f"[DEBUG] Found {len(children)} children in folder '{folder_path}' ({item_id})")
                else:
                    print(f"[DEBUG] Found {len(children)} children in folder ({item_id})")
            return children
        else:
            raise Exception(f"List children failed: {children_response.status_code} - {children_response.text}")

    except Exception as e:
        print(f"[!] Error listing folder children: {str(e)}")
        if is_debug_metadata_enabled():
            import traceback
            print(f"[DEBUG] Traceback: {traceback.format_exc()}")
        return None


def batch_update_filehash_fields(site_url, list_name, updates_list,
                                 tenant_id, client_id, client_secret,
                                 login_endpoint, graph_endpoint, batch_size=20,
                                 requery_item_ids=False):
    """
    Update multiple FileHash fields in SharePoint using batch requests.

    Args:
        site_url (str): Full SharePoint site URL
        list_name (str): Name of the document library
        updates_list (list): List of tuples in one of two formats:
            - Normal mode: (item_id, filename, hash_value, display_path)
            - Requery mode: (parent_item_id, filename, None, hash_value, is_update, display_path)
        tenant_id (str): Azure AD tenant ID
        client_id (str): App registration client ID
        client_secret (str): App registration client secret
        login_endpoint (str): Azure AD endpoint
        graph_endpoint (str): Graph API endpoint
        batch_size (int): Items per batch request (max 20 for Graph API)
        requery_item_ids (bool): If True, re-query item IDs using parent_id + filename

    Returns:
        dict: Mapping of {item_id: success_bool} or {index: success_bool} for requery mode

    Example:
        # Normal mode (first attempt)
        updates = [('item1', 'file1.txt', 'hash1', 'docs/file1.txt')]
        results = batch_update_filehash_fields(site_url, lib, updates, ...)

        # Requery mode (retry)
        updates = [('parent1', 'file1.txt', None, 'hash1', True, 'docs/file1.txt')]
        results = batch_update_filehash_fields(..., requery_item_ids=True)
    """
    try:
        if not updates_list:
            return {}

        # Get token for Graph API
        token = acquire_token(tenant_id, client_id, client_secret, login_endpoint, graph_endpoint)

        if 'access_token' not in token:
            print(f"[!] Failed to acquire token for batch updates")
            return {(parent_id, filename): False for parent_id, filename, _, _, _ in updates_list}

        headers = {
            'Authorization': f"Bearer {token['access_token']}",
            'Content-Type': 'application/json'
        }

        # Parse site URL to get site ID
        site_parts = site_url.replace('https://', '').split('/')
        host_name = site_parts[0]
        site_name = site_parts[2] if len(site_parts) > 2 else ''

        # Get site ID
        site_endpoint = f"https://{graph_endpoint}/v1.0/sites/{host_name}:/sites/{site_name}"
        site_response = make_graph_request_with_retry(site_endpoint, headers, method='GET')

        if site_response.status_code != 200:
            print(f"[!] Failed to get site information for batch updates")
            return {(parent_id, filename): False for parent_id, filename, _, _, _ in updates_list}

        site_data = site_response.json()
        site_id = site_data.get('id')

        if not site_id:
            print("[!] Could not retrieve site ID for batch updates")
            return {(parent_id, filename): False for parent_id, filename, _, _, _ in updates_list}

        # Get list ID and drive ID
        lists_endpoint = f"https://{graph_endpoint}/v1.0/sites/{site_id}/lists"
        lists_response = make_graph_request_with_retry(lists_endpoint, headers, method='GET')

        if lists_response.status_code != 200:
            print(f"[!] Failed to get lists for batch updates")
            return {(parent_id, filename): False for parent_id, filename, _, _, _ in updates_list}

        lists_data = lists_response.json()
        list_id = None
        drive_id = None

        for sp_list in lists_data.get('value', []):
            if sp_list.get('displayName') == list_name or sp_list.get('name') == list_name:
                list_id = sp_list.get('id')
                # Document libraries have an associated drive
                drive_info = sp_list.get('list', {}).get('template')
                if drive_info == 'documentLibrary' or 'drive' in sp_list:
                    # Get the drive from the list's drive property
                    drive_id = sp_list.get('id')  # For document libraries, list ID = drive ID
                break

        if not list_id:
            print(f"[!] Could not find list '{list_name}' for batch updates")
            return {(parent_id, filename): False for parent_id, filename, _, _, _ in updates_list}

        # For document libraries, we need to get the actual drive ID
        # Query the drive directly using the site and list
        drives_endpoint = f"https://{graph_endpoint}/v1.0/sites/{site_id}/drives"
        drives_response = make_graph_request_with_retry(drives_endpoint, headers, method='GET')

        if drives_response.status_code == 200:
            drives_data = drives_response.json()
            for drive in drives_data.get('value', []):
                # Match drive to our list - drives for document libraries have the same name
                if drive.get('name') == list_name:
                    drive_id = drive.get('id')
                    break

        # Handle re-query mode vs normal mode
        if requery_item_ids:
            # Requery mode: Query fresh item IDs for failed files
            if is_debug_enabled():
                print(f"[DEBUG] Re-querying list item IDs for {len(updates_list)} files...")

            import urllib.parse
            item_id_map = {}

            for idx, item in enumerate(updates_list):
                parent_id, filename = item[0], item[1]
                try:
                    encoded_filename = urllib.parse.quote(filename)
                    item_url = f"https://{graph_endpoint}/v1.0/sites/{site_id}/drives/{drive_id}/items/{parent_id}:/{encoded_filename}?$expand=listItem"

                    response = make_graph_request_with_retry(item_url, headers, method='GET')

                    if response.status_code == 200:
                        data = response.json()
                        list_item = data.get('listItem')
                        if list_item and 'id' in list_item:
                            item_id_map[idx] = list_item['id']
                except Exception:
                    pass  # Will be marked as failed below

        # Process updates in batches
        results = {}
        total_batches = (len(updates_list) + batch_size - 1) // batch_size

        for batch_num in range(0, len(updates_list), batch_size):
            batch = updates_list[batch_num:batch_num+batch_size]
            batch_index = batch_num // batch_size + 1

            # Build JSON batch request
            batch_request = {"requests": []}

            for idx, item in enumerate(batch):
                global_idx = batch_num + idx

                if requery_item_ids:
                    # Requery mode: (parent_id, filename, None, hash, is_update, display_path)
                    list_item_id = item_id_map.get(global_idx)
                    hash_value = item[3]
                    display_path = item[5]
                else:
                    # Normal mode: (item_id, filename, hash, display_path)
                    list_item_id = item[0]
                    hash_value = item[2]
                    display_path = item[3]

                if not list_item_id:
                    results[global_idx if requery_item_ids else list_item_id] = False
                    continue

                request_item = {
                    "id": str(idx),
                    "method": "PATCH",
                    "url": f"/sites/{site_id}/lists/{list_id}/items/{list_item_id}/fields",
                    "body": {"FileHash": hash_value},
                    "headers": {"Content-Type": "application/json"}
                }
                batch_request["requests"].append(request_item)

            # Skip if no requests in batch
            if not batch_request["requests"]:
                continue

            # Send batch request
            batch_endpoint = f"https://{graph_endpoint}/v1.0/$batch"

            try:
                batch_response = make_graph_request_with_retry(
                    batch_endpoint,
                    headers,
                    method='POST',
                    json_data=batch_request
                )

                if batch_response.status_code == 200:
                    batch_data = batch_response.json()
                    batch_results = batch_data.get('responses', [])

                    for result in batch_results:
                        try:
                            request_id = int(result['id'])
                            global_idx = batch_num + request_id
                            item = batch[request_id]

                            if requery_item_ids:
                                key = global_idx
                                list_item_id = item_id_map.get(global_idx)
                                display_path = item[5]
                                filename = item[1]
                            else:
                                list_item_id = item[0]
                                key = list_item_id
                                display_path = item[3]
                                filename = item[1]

                            success = 200 <= result['status'] < 300
                            results[key] = success

                            # Show individual file success/failure
                            if success:
                                if is_debug_enabled():
                                    print(f"[DEBUG] ✓ Updated FileHash for {display_path} ({filename})")
                            else:
                                print(f"[DEBUG] × Failed to update FileHash for {display_path} ({filename}): HTTP {result.get('status')}")

                        except Exception:
                            continue

                    # Mark missing items as failed
                    for idx in range(len(batch)):
                        global_idx = batch_num + idx
                        if requery_item_ids:
                            key = global_idx
                        else:
                            key = batch[idx][0]

                        if key not in results:
                            results[key] = False

                else:
                    # Entire batch failed
                    for idx in range(len(batch)):
                        global_idx = batch_num + idx
                        key = global_idx if requery_item_ids else batch[idx][0]
                        results[key] = False

            except Exception as batch_error:
                print(f"[!] Error processing batch {batch_index}: {str(batch_error)[:200]}")
                for idx in range(len(batch)):
                    global_idx = batch_num + idx
                    key = global_idx if requery_item_ids else batch[idx][0]
                    results[key] = False

        return results

    except Exception as e:
        print(f"[!] Batch update failed: {str(e)[:400]}")
        # Return all failed
        if requery_item_ids:
            return {idx: False for idx in range(len(updates_list))}
        else:
            return {item[0]: False for item in updates_list}
