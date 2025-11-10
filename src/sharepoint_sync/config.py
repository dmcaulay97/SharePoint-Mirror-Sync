# -*- coding: utf-8 -*-
"""
Configuration management for SharePoint sync.

This module handles command-line argument parsing and configuration setup.
"""

import sys


class Config:
    """Configuration for SharePoint sync operations"""

    def __init__(self):
        """
        Parse command-line arguments and initialize configuration.

        Arguments are parsed from sys.argv in the following order:
        1. site_name - SharePoint site name
        2. sharepoint_host_name - SharePoint domain
        3. tenant_id - Azure AD tenant ID
        4. client_id - App registration client ID
        5. client_secret - App registration client secret
        6. upload_path - Target path in SharePoint
        7. file_path - Local file/glob pattern to upload
        8. max_retry (optional) - Max retry attempts (default: 3)
        9. login_endpoint (optional) - Azure AD endpoint (default: login.microsoftonline.com)
        10. graph_endpoint (optional) - Graph API endpoint (default: graph.microsoft.com)
        11. recursive (optional) - Enable recursive glob (default: False)
        12. force_upload (optional) - Force upload all files (default: False)
        13. convert_md_to_html (optional) - Convert markdown to HTML (default: True)
        14. force_md_to_html_regeneration (optional) - Force regenerate HTML from .md even if unchanged (default: False)
        15. exclude_patterns (optional) - Comma-separated exclusion patterns (default: "")
        16. sync_delete (optional) - Delete SharePoint files not in sync set (default: False)
        17. sync_delete_whatif (optional) - Preview deletions without actually deleting (default: True)
        18. max_upload_workers (optional) - Max concurrent uploads (default: 4, respects Graph API limits)
        19. debug (optional) - Enable general debug output (default: False)
        20. debug_metadata (optional) - Enable metadata-specific debug output (default: False)
        """
        # Required arguments
        self.site_name = sys.argv[1]
        self.sharepoint_host_name = sys.argv[2]
        self.tenant_id = sys.argv[3]
        self.client_id = sys.argv[4]
        self.client_secret = sys.argv[5]
        self.upload_path = sys.argv[6]
        self.file_path = sys.argv[7]

        # Optional arguments with defaults
        self.max_retry = int(sys.argv[8]) if len(sys.argv) > 8 and sys.argv[8] else 3
        self.login_endpoint = sys.argv[9] if len(sys.argv) > 9 and sys.argv[9] else "login.microsoftonline.com"
        self.graph_endpoint = sys.argv[10] if len(sys.argv) > 10 and sys.argv[10] else "graph.microsoft.com"
        self.recursive = (sys.argv[11] if len(sys.argv) > 11 else "false").lower() == "true"
        self.force_upload = (sys.argv[12] if len(sys.argv) > 12 else "false").lower() == "true"
        self.convert_md_to_html = (sys.argv[13] if len(sys.argv) > 13 else "true").lower() == "true"
        self.force_md_to_html_regeneration = (sys.argv[14] if len(sys.argv) > 14 else "false").lower() == "true"
        self.exclude_patterns = sys.argv[15] if len(sys.argv) > 15 and sys.argv[15] else ""
        self.sync_delete = (sys.argv[16] if len(sys.argv) > 16 else "false").lower() == "true"
        self.sync_delete_whatif = (sys.argv[17] if len(sys.argv) > 17 else "true").lower() == "true"

        # Parallel processing configuration (auto-detect optimal values)
        import os as os_module
        cpu_count = os_module.cpu_count() or 4

        # Max upload workers: Default 4 (Graph API concurrent request limit)
        # Can be overridden but should not exceed 10 to respect API limits
        # WARNING: Starting September 30, 2025, Microsoft will reduce per-app/per-user
        # throttling limits to HALF the total per-tenant limit. Monitor for increased
        # 429 responses after this date. Default of 4 workers should remain safe.
        if len(sys.argv) > 18 and sys.argv[18]:
            self.max_upload_workers = min(int(sys.argv[18]), 10)
        else:
            self.max_upload_workers = 4  # Safe default for Graph API

        # Max markdown workers: Default 4 (mermaid-cli subprocess limit)
        # Balance between parallelism and Chromium memory usage
        self.max_markdown_workers = min(4, cpu_count)

        # Debug flags
        self.debug = (sys.argv[19] if len(sys.argv) > 19 else "false").lower() == "true"
        self.debug_metadata = (sys.argv[20] if len(sys.argv) > 20 else "false").lower() == "true"

        # Specify upload library flag
        self.specify_upload_library = (sys.argv[21] if len(sys.argv) > 21 else "false").lower() == "true"

        # Derived values
        self.tenant_url = f'https://{self.sharepoint_host_name}/sites/{self.site_name}'
        self.exclude_patterns_list = [p.strip() for p in self.exclude_patterns.split(',') if p.strip()]

    def validate(self):
        """
        Validate configuration values.

        Raises:
            ValueError: If configuration is invalid
        """
        if not self.site_name:
            raise ValueError("site_name cannot be empty")
        if not self.sharepoint_host_name:
            raise ValueError("sharepoint_host_name cannot be empty")
        if not self.tenant_id:
            raise ValueError("tenant_id cannot be empty")
        if not self.client_id:
            raise ValueError("client_id cannot be empty")
        if not self.client_secret:
            raise ValueError("client_secret cannot be empty")
        if not self.upload_path:
            raise ValueError("upload_path cannot be empty")
        if not self.file_path:
            raise ValueError("file_path cannot be empty")
        if self.max_retry < 0:
            raise ValueError("max_retry must be non-negative")


def parse_config():
    """
    Parse configuration from command-line arguments.

    Returns:
        Config: Configured Config object

    Raises:
        ValueError: If configuration is invalid
        IndexError: If required arguments are missing
    """
    config = Config()
    config.validate()
    return config
