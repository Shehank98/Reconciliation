import pandas as pd
import re



def clean_column_names(df):
    """Clean up column names by removing _x and _y suffixes and handling duplicates."""
    # Make a copy of the DataFrame to avoid modifying the original
    df = df.copy()
    
    # Initialize mappings and tracking
    col_mapping = {}
    seen_base_cols = {}
    
    for col in df.columns:
        base_name = col
        # Handle _x and _y suffixes
        if col.endswith(('_x', '_y')):
            base_name = col[:-2]
        
        # Track occurrences of base names
        if base_name not in seen_base_cols:
            seen_base_cols[base_name] = 1
            col_mapping[col] = base_name
        else:
            # For duplicates, append a number
            seen_base_cols[base_name] += 1
            col_mapping[col] = f"{base_name}_{seen_base_cols[base_name]}"
    
    # Rename columns
    df = df.rename(columns=col_mapping)
    
    # Remove any remaining duplicates by keeping first occurrence
    df = df.loc[:, ~df.columns.duplicated(keep='first')]
    
    return df

def standardize_dataframe(df, date_col=None, time_col=None, theme_col=None, program_col=None):
    """Standardize DataFrame columns and formats for consistent processing."""
    # Create a copy and reset index
    result_df = df.copy()
    result_df.reset_index(drop=True, inplace=True)
    
    # First handle duplicates at the column level
    result_df = result_df.loc[:, ~result_df.columns.duplicated(keep='first')]
    
    # Clean column names initially
    result_df = clean_column_names(result_df)
    print("Columns after initial cleaning:", result_df.columns.tolist())    # Auto-detect time column with improved handling of Advt_time vs Prog_time
    if not time_col or time_col not in result_df.columns:
        # First try to find existing Advt_time (prioritize this exact column)
        time_col = next(
            (col for col in result_df.columns if col == 'Advt_time'),
            next(
                (col for col in result_df.columns if col == 'Advt_Time'),
                None
            )
        )
        if not time_col:
            # Only look for other time columns if Advt_time doesn't exist
            time_col = next(
                (col for col in result_df.columns 
                 if col not in ['Advt_time', 'prog_time'] and any(term in col.lower() for term in ['time', 'spot', 'air', 'clock'])),
                None
            )
        print(f"Auto-detected time column: {time_col}")
    if not theme_col:
        pass
    if not program_col:
        pass
    # Build rename dictionary with better debug output and improved time column handling
    rename_dict = {}
    print(f"\nInitial columns: {result_df.columns.tolist()}")
    print(f"Detected theme column: {theme_col}")
    print(f"Detected program column: {program_col}")
    print(f"Detected time column: {time_col}")
    
    # Check if we already have either Advt_time or Advt_Time
    has_advt_time = 'Advt_time' in result_df.columns or 'Advt_Time' in result_df.columns
    
    if theme_col and theme_col in result_df.columns and theme_col.lower() != 'advt_theme':
        pass
    if program_col and program_col in result_df.columns and program_col.lower() != 'program':
        pass
    if not has_advt_time and time_col and time_col in result_df.columns and time_col.lower() != 'advt_time':
        pass

    # Apply column renames
    if rename_dict:
        result_df = result_df.rename(columns=rename_dict)
        result_df = clean_column_names(result_df)

    # Ensure we have a proper Advt_Theme column
    if 'Advt_Theme' not in result_df.columns and theme_col and theme_col in result_df.columns:
        result_df['Advt_Theme'] = result_df[theme_col]

    # Create Normalized_Theme from Advt_Theme
    if 'Advt_Theme' in result_df.columns:
        result_df['Normalized_Theme'] = result_df['Advt_Theme'].apply(normalize_theme_name)    # Handle time column - improved handling and debug output    print("\nHandling time column...")
    existing_time_col = None
    
    # First verify if we already have Advt_time or Advt_Time
    if 'Advt_time' in result_df.columns:
        existing_time_col = 'Advt_time'
        print("Found existing Advt_time column")
    elif 'Advt_Time' in result_df.columns:
        existing_time_col = 'Advt_Time'
        print("Found existing Advt_Time column")
    elif time_col and time_col in result_df.columns:
        # Only use detected time_col if we don't already have Advt_time
        if time_col != 'Prog_time':  # Never use Prog_time
            existing_time_col = time_col
            print(f"Using detected time column: {time_col}")
    
    if existing_time_col:
        if existing_time_col.lower() != 'advt_time':
            # Create new standardized column
            print(f"Creating standardized time column from {existing_time_col}")
            result_df['Advt_Time'] = result_df[existing_time_col].astype(str).apply(preprocess_time_format)
            # Drop the old column to avoid confusion
            if existing_time_col != 'Advt_Time':
                result_df = result_df.drop(columns=[existing_time_col])
                print(f"Dropped original time column: {existing_time_col}")
        else:
            # Just standardize the format of existing column
            print("Standardizing existing Advt_Time column format")
            result_df['Advt_Time'] = result_df[existing_time_col].astype(str).apply(preprocess_time_format)
        
        # Calculate seconds for matching
        result_df['Time_Seconds'] = result_df['Advt_Time'].astype(str).apply(time_to_seconds)
        print("Time values converted to seconds for matching")
    else:
        # No appropriate time column found
        print(f"Warning: No time column found. Available columns: {result_df.columns.tolist()}")
        print("Using default time values")
        result_df['Advt_Time'] = '00:00:00'
        result_df['Time_Seconds'] = 0
    
    print(f"Final time column name: Advt_Time")

    # Handle date column
    if date_col and date_col in df.columns:
        try:
            result_df['Date'] = pd.to_datetime(df[date_col], errors='coerce')
            result_df['Dd'] = result_df['Date'].dt.day
            result_df['Mn'] = result_df['Date'].dt.month
            result_df['Yr'] = result_df['Date'].dt.year
        except Exception as e:
            print(f"Error processing date column: {e}")

    # Final cleanup and column verification
    result_df = clean_column_names(result_df)
    result_df = result_df.loc[:, ~result_df.columns.duplicated(keep='first')]
    
    # Verify and create required columns
    required_columns = ['Advt_Theme', 'Program', 'Advt_time', 'Normalized_Theme']
    for col in required_columns:
        if col not in result_df.columns:
            if col == 'Advt_time':
                result_df[col] = '00:00:00'
            elif col == 'Program':
                result_df[col] = 'Unknown Program' 
            elif col == 'Advt_Theme':
                result_df[col] = 'Unknown Theme'
            elif col == 'Normalized_Theme':
                if 'Advt_Theme' in result_df.columns:
                    result_df[col] = result_df['Advt_Theme'].apply(normalize_theme_name)
                else:
                    result_df[col] = 'unknown theme'

    return result_df.copy()

def preprocess_time_format(time_value):
    """Extract and standardize time format from various inputs."""
    try:
        if pd.isna(time_value):
            return '00:00:00'
            
        if isinstance(time_value, str):
            time_pattern = re.search(r'(\d{1,2}:\d{2}(?::\d{2})?)', str(time_value))
            if time_pattern:
                time_str = time_pattern.group(1)
                if time_str.count(':') == 1:
                    time_str += ':00'
                return time_str
            return '00:00:00'
        elif hasattr(time_value, 'strftime'):
            return time_value.strftime('%H:%M:%S')
        return '00:00:00'
    except Exception:
        return '00:00:00'

def time_to_seconds(time_str):
    """Convert time string (HH:MM:SS) to seconds."""
    try:
        if pd.isna(time_str):
            return 0
            
        time_str = str(time_str)
        time_parts = time_str.split(':')
        
        if len(time_parts) >= 3:
            return int(time_parts[0]) * 3600 + int(time_parts[1]) * 60 + int(time_parts[2])
        elif len(time_parts) == 2:
            return int(time_parts[0]) * 3600 + int(time_parts[1]) * 60
        return 0
        
    except (ValueError, IndexError):
        return 0

def normalize_theme_name(theme):
    """
    Normalize theme name for consistent comparison by standardizing format.
    
    Args:
        theme: The theme name to normalize
        
    Returns:
        str: Normalized theme name
    """
    if pd.isna(theme):
        return ""
        
    # Convert to string and clean
    theme = str(theme).strip()
    
    # Convert to uppercase for consistent handling
    theme = theme.upper()
    
    # Remove common prefix/suffix variations with case-insensitive matching
    prefixes = ['THEME:', 'THEME-', 'THEME_', 'TH:', 'TH-', 'TH_']
    suffixes = ['-THEME', '_THEME', ' THEME', '-TH', '_TH', ' TH']
    
    for prefix in prefixes:
        if theme.startswith(prefix):
            theme = theme[len(prefix):]
            break
            
    for suffix in suffixes:
        if theme.endswith(suffix):
            theme = theme[:-len(suffix)]
            break
    
    # Replace special characters with spaces
    theme = re.sub(r'[^\w\s-]', ' ', theme)
    
    # Remove duration patterns (e.g., "30S", "10 SEC", etc.)
    theme = re.sub(r'\b\d+\s*(?:SEC|SECS|S)?\b', '', theme)
    
    # Normalize whitespace
    theme = ' '.join(theme.split())
    
    # Strip remaining whitespace and ensure non-empty
    theme = theme.strip()
    
    return theme if theme else ""