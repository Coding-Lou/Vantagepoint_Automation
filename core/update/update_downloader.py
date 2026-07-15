"""
Update downloader with progress callback support.
"""
from typing import Optional, Callable
from pathlib import Path
import requests
import os
import shutil


class UpdateDownloader:
    """
    Download update files with progress tracking.
    """
    
    def __init__(self, progress_callback: Optional[Callable[[int, int], None]] = None):
        """
        Initialize downloader.
        
        Args:
            progress_callback: Optional callback function(current_bytes, total_bytes)
        """
        self.progress_callback = progress_callback
    
    def download(
        self,
        download_url: str,
        save_path: Path,
        timeout: int = 300
    ) -> bool:
        """
        Download file from URL to specified path.
        
        Args:
            download_url: URL to download from
            save_path: Path to save the downloaded file
            timeout: Request timeout in seconds
            
        Returns:
            True if download successful, False otherwise
            
        Raises:
            requests.RequestException: On network errors
            OSError: On file system errors (permissions, disk space)
        """
        try:
            # Ensure parent directory exists
            save_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Check available disk space (cross-platform)
            try:
                disk_usage = shutil.disk_usage(save_path.parent)
                free_space = disk_usage.free
            except Exception:
                # If disk space check fails, continue anyway
                free_space = float('inf')
            
            # Start download with streaming
            response = requests.get(download_url, stream=True, timeout=timeout)
            response.raise_for_status()
            
            total_size = int(response.headers.get('content-length', 0))
            downloaded = 0
            chunk_size = 8192  # 8KB chunks
            
            # Check disk space before writing
            if total_size > 0 and free_space < total_size * 1.1:  # 10% buffer
                raise OSError("Insufficient disk space for download")
            
            # Download and write file
            with open(save_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=chunk_size):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        
                        # Call progress callback if provided
                        if self.progress_callback:
                            try:
                                self.progress_callback(downloaded, total_size)
                            except Exception:
                                # Don't let callback errors break download
                                pass
                        
                        # Check disk space during download
                        if total_size > 0 and downloaded < total_size:
                            try:
                                disk_usage = shutil.disk_usage(save_path.parent)
                                free_space = disk_usage.free
                                if free_space < (total_size - downloaded) * 1.1:
                                    raise OSError("Insufficient disk space during download")
                            except Exception:
                                # If disk space check fails, continue anyway
                                pass
            
            # Verify download completed
            if total_size > 0 and downloaded != total_size:
                raise IOError(f"Download incomplete: {downloaded}/{total_size} bytes")
            
            return True
            
        except requests.RequestException as e:
            # Network errors
            raise
        except OSError as e:
            # File system errors (permissions, disk space)
            raise
        except Exception as e:
            # Other errors
            raise IOError(f"Download failed: {str(e)}") from e
