#!/usr/bin/env python3
"""
Script to delete all non-.docx files from the data directory and its subdirectories.
Includes a confirmation prompt before deletion for safety.
"""

import os
import sys
from pathlib import Path

def list_non_docx_files(directory):
    """
    List all non-.docx files in the given directory and its subdirectories.
    
    Args:
        directory (str): Path to the directory to search
        
    Returns:
        list: List of non-.docx file paths
    """
    non_docx_files = []
    
    for root, dirs, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            if not file.lower().endswith('.docx'):
                non_docx_files.append(file_path)
    
    return non_docx_files

def delete_files(file_list):
    """
    Delete the files in the given list.
    
    Args:
        file_list (list): List of file paths to delete
        
    Returns:
        tuple: (number of files deleted successfully, list of files that failed to delete)
    """
    deleted_count = 0
    failed_files = []
    
    for file_path in file_list:
        try:
            os.remove(file_path)
            print(f"Deleted: {file_path}")
            deleted_count += 1
        except Exception as e:
            print(f"Failed to delete {file_path}: {e}")
            failed_files.append(file_path)
    
    return deleted_count, failed_files

def main():
    # Define the data directory path
    data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
    
    # Check if the data directory exists
    if not os.path.isdir(data_dir):
        print(f"Error: Data directory not found at {data_dir}")
        sys.exit(1)
    
    # List all non-.docx files
    non_docx_files = list_non_docx_files(data_dir)
    
    if not non_docx_files:
        print("No non-.docx files found. Nothing to delete.")
        return
    
    # Show the files that will be deleted
    print(f"Found {len(non_docx_files)} non-.docx files to delete:")
    for file_path in non_docx_files:
        print(f"  {file_path}")
    
    # Ask for confirmation
    confirmation = input("\nDo you want to delete these files? (yes/no): ")
    
    if confirmation.lower() not in ['yes', 'y']:
        print("Operation cancelled.")
        return
    
    # Delete the files
    deleted_count, failed_files = delete_files(non_docx_files)
    
    # Report results
    print(f"\nDeletion complete. {deleted_count} files deleted.")
    
    if failed_files:
        print(f"Failed to delete {len(failed_files)} files:")
        for file_path in failed_files:
            print(f"  {file_path}")

if __name__ == "__main__":
    main()
