"""Module for combining multiple questionnaires into a single input."""

import pandas as pd
from pathlib import Path
from typing import List, Union


def load_questionnaire(file_path: Union[str, Path]) -> pd.DataFrame:
    """
    Load a questionnaire from an Excel file.
    
    Args:
        file_path: Path to the Excel file
        
    Returns:
        DataFrame containing the questionnaire data
    """
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        raise ValueError(f"Error loading file {file_path}: {str(e)}")


def combine_questionnaires(file_paths: List[Union[str, Path]]) -> pd.DataFrame:
    """
    Combine multiple questionnaire files into a single DataFrame.
    
    Args:
        file_paths: List of paths to Excel files
        
    Returns:
        Combined DataFrame from all questionnaires
    """
    combined_data = []
    
    for file_path in file_paths:
        df = load_questionnaire(file_path)
        combined_data.append(df)
    
    # Concatenate all questionnaires
    result = pd.concat(combined_data, ignore_index=True)
    return result


def save_combined_questionnaire(df: pd.DataFrame, output_path: Union[str, Path]) -> None:
    """
    Save combined questionnaire to file.
    
    Args:
        df: DataFrame to save
        output_path: Path to save the file
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    if output_path.suffix.lower() == '.xlsx':
        df.to_excel(output_path, index=False)
    elif output_path.suffix.lower() == '.csv':
        df.to_csv(output_path, index=False)
    else:
        raise ValueError("Unsupported file format. Use .xlsx or .csv")


if __name__ == "__main__":
    print("Questionnaire combination module loaded")
