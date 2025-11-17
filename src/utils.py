"""Utility functions for the project."""

import json
from pathlib import Path
from typing import Any, Dict


def save_json(data: Dict[str, Any], file_path: str) -> None:
    """
    Save data to a JSON file.
    
    Args:
        data: Dictionary to save
        file_path: Path to save the file
    """
    Path(file_path).parent.mkdir(parents=True, exist_ok=True)
    with open(file_path, 'w') as f:
        json.dump(data, f, indent=2)


def load_json(file_path: str) -> Dict[str, Any]:
    """
    Load data from a JSON file.
    
    Args:
        file_path: Path to the JSON file
        
    Returns:
        Loaded dictionary
    """
    with open(file_path, 'r') as f:
        return json.load(f)


def format_text(text: str, max_length: int = 500) -> str:
    """
    Format text for display or API submission.
    
    Args:
        text: Text to format
        max_length: Maximum length before truncation
        
    Returns:
        Formatted text
    """
    text = text.strip()
    if len(text) > max_length:
        return text[:max_length] + "..."
    return text


if __name__ == "__main__":
    print("Utilities module loaded")
