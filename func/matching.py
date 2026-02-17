import pandas as pd
from func.format import normalize_theme_name

def match_media_watch_with_tc(media_watch_df, tc_df, theme_mapping, ignore_date=False, time_tolerance=30):
    """
    Match Media Watch data with TC data based on theme mapping.
    ENHANCED: Supports multiple TC themes mapping to same LMRB theme.

    Parameters:
        media_watch_df (pd.DataFrame): The Media Watch DataFrame.
        tc_df (pd.DataFrame): The TC DataFrame.
        theme_mapping (list of dict): A list of dictionaries containing theme mappings.
        ignore_date (bool): Whether to ignore date in matching.
        time_tolerance (int): Time tolerance in seconds for matching.

    Returns:
        matched_df (pd.DataFrame): DataFrame containing matched rows.
        matched_mw_indices (list): List of indices from Media Watch that were matched.
        program_mismatched_df (pd.DataFrame): DataFrame containing time-matched but program-mismatched rows.
        matched_tc_keys (list): List of unique keys for matched TC rows.
    """

    # Initialize filter stats for tracking match process
    filter_stats = {
        "total_media_watch_rows": len(media_watch_df),
        "no_normalized_mapping": 0,
        "no_matches_after_theme_filter": 0,
        "no_matches_after_date_filter": 0,
        "no_matches_after_time_filter": 0,
        "matches_found": 0,
        "program_mismatched_but_time_matched": 0,
        "multiple_mapping_attempts": 0,  # NEW: Track multiple mapping usage
        "successful_fallback_matches": 0,  # NEW: Track successful fallback to 2nd/3rd mapping
        "mapping_attempts_breakdown": {}  # NEW: Track which mapping number succeeded
    }

    print("\n" + "="*80)
    print("🚀 ENHANCED MATCHING: Multiple TC Mappings per LMRB Theme")
    print("="*80)
    print(f"Total Media Watch rows: {filter_stats['total_media_watch_rows']}")

    # Validate input DataFrames
    if media_watch_df.empty or tc_df.empty:
        print("❌ One or both input DataFrames are empty.")
        return pd.DataFrame(), [], pd.DataFrame(), []

    # Make copies to avoid modifying originals
    media_watch_df = media_watch_df.copy().reset_index(drop=True)  
    tc_df = tc_df.copy().reset_index(drop=True)

    # Print column information
    print(f"\n📊 Data Info:")
    print(f"Media Watch columns: {media_watch_df.columns.tolist()}")
    print(f"TC columns: {tc_df.columns.tolist()}")

    # Ensure Normalized_Theme and Duration exist in both dataframes
    for df, name in [(media_watch_df, "Media Watch"), (tc_df, "TC")]:
        if 'Normalized_Theme' not in df.columns and 'Advt_Theme' in df.columns:
            df['Normalized_Theme'] = df['Advt_Theme'].apply(normalize_theme_name)
        # Ensure duration column exists (try to find it)
        if 'Duration' not in df.columns:
            dur_col = next((col for col in df.columns if 'dur' in col.lower()), None)
            if dur_col:
                df['Duration'] = df[dur_col]
            else:
                df['Duration'] = None

    # ============================================================
    # ENHANCED: Create multiple mapping dictionary structure
    # ============================================================
    print(f"\n🗺️ Processing theme mappings...")
    mw_to_tc_multiple = {}  # {(mw_theme, mw_dur): [(tc_theme1, tc_dur1), (tc_theme2, tc_dur2), ...]}

    for mapping in theme_mapping:
        try:
            # Get mapping keys with fallback
            mw_theme_key = 'media_watch_theme' if 'media_watch_theme' in mapping else 'lmrb_theme'
            mw_dur_key = 'lmrb_duration' if 'lmrb_duration' in mapping else 'lmrb_dur'
            tc_theme_key = 'tc_theme'
            tc_dur_key = 'tc_duration' if 'tc_duration' in mapping else 'tc_dur'
            
            mw_theme = mapping.get(mw_theme_key)
            mw_dur = str(mapping.get(mw_dur_key, '')).strip()
            tc_theme = mapping.get(tc_theme_key)
            tc_dur = str(mapping.get(tc_dur_key, '')).strip()
            
            # Skip if essential data is missing
            if not mw_theme or not tc_theme or not mw_dur or not tc_dur:
                continue
            
            mw_theme_norm = normalize_theme_name(mw_theme)
            tc_theme_norm = normalize_theme_name(tc_theme)
            
            # Create key for LMRB theme+duration
            mw_key = (mw_theme_norm, mw_dur)
            tc_mapping = (tc_theme_norm, tc_dur)
            
            # Store multiple TC mappings for same LMRB theme
            if mw_key not in mw_to_tc_multiple:
                mw_to_tc_multiple[mw_key] = []
            mw_to_tc_multiple[mw_key].append(tc_mapping)
            
        except Exception as e:
            print(f"⚠️ Error processing mapping: {e}")
            continue

    if not mw_to_tc_multiple:
        print("❌ Error: No valid theme mappings found!")
        return pd.DataFrame(), [], pd.DataFrame(), []

    # Print enhanced mapping info
    total_lmrb_combinations = len(mw_to_tc_multiple)
    total_tc_mappings = sum(len(tc_list) for tc_list in mw_to_tc_multiple.values())
    multiple_mapping_count = len([tc_list for tc_list in mw_to_tc_multiple.values() if len(tc_list) > 1])

    print(f"\n📋 Mapping Summary:")
    print(f"   LMRB theme combinations: {total_lmrb_combinations}")
    print(f"   Total TC mappings: {total_tc_mappings}")
    print(f"   LMRB themes with multiple TC mappings: {multiple_mapping_count}")

    print(f"\n🔀 Enhanced Theme Mapping Structure:")
    for (mw_theme, mw_dur), tc_list in mw_to_tc_multiple.items():
        if len(tc_list) > 1:
            print(f"   🔀 {mw_theme} ({mw_dur}s) → {len(tc_list)} TC options:")
            for i, (tc_theme, tc_dur) in enumerate(tc_list, 1):
                print(f"      {i}. {tc_theme} ({tc_dur}s)")
        else:
            tc_theme, tc_dur = tc_list[0]
            print(f"   📍 {mw_theme} ({mw_dur}s) → {tc_theme} ({tc_dur}s)")

    # Initialize results
    matched_results = []
    matched_mw_indices = []  
    program_mismatches = []
    matched_tc_indices = []
    already_matched_tc_indices = set()

    # ============================================================
    # ENHANCED: Multiple mapping matching loop
    # ============================================================
    print(f"\n🔄 Starting enhanced matching process...")
    total_rows = len(media_watch_df)

    for i, mw_row in enumerate(media_watch_df.itertuples(), 1):
        # Progress indicator
        if i % 100 == 0 or i == total_rows:
            print(f"\r🔄 Processing: {i}/{total_rows} ({i/total_rows*100:.1f}%)", end='')
        
        # Skip rows with missing required data
        if pd.isna(mw_row.Advt_Theme) or pd.isna(mw_row.Program) or pd.isna(mw_row.Advt_time):
            continue

        mw_theme_norm = normalize_theme_name(mw_row.Advt_Theme)
        mw_dur = str(getattr(mw_row, 'Duration', '')).strip()
        
        # Check if theme has any mappings
        mw_key = (mw_theme_norm, mw_dur)
        if mw_key not in mw_to_tc_multiple:
            filter_stats["no_normalized_mapping"] += 1
            continue

        # Get all TC mappings for this LMRB theme
        tc_mappings_list = mw_to_tc_multiple[mw_key]
        
        # Track if we're using multiple mappings
        if len(tc_mappings_list) > 1:
            filter_stats["multiple_mapping_attempts"] += 1

        # Try each TC mapping until we find a match
        match_found = False
        best_match = None
        match_attempt = 0
        successful_tc_theme = None
        
        for tc_theme_norm, tc_dur in tc_mappings_list:
            match_attempt += 1
            
            # Find TC matches for this specific TC theme
            tc_matches = tc_df[
                (tc_df['Normalized_Theme'] == tc_theme_norm) &
                (tc_df['Duration'].astype(str).str.strip() == tc_dur)
            ].copy()
            
            if tc_matches.empty:
                continue  # Try next TC mapping
            
            # Apply date filter unless ignored
            if not ignore_date:
                if all(hasattr(mw_row, attr) for attr in ['Dd', 'Mn', 'Yr']):
                    date_filter = (
                        (tc_matches['Dd'] == mw_row.Dd) &
                        (tc_matches['Mn'] == mw_row.Mn) &
                        (tc_matches['Yr'] == mw_row.Yr)
                    )
                    tc_matches = tc_matches[date_filter]
            
            if tc_matches.empty:
                continue  # Try next TC mapping
                
            # Calculate time differences
            mw_time_seconds = getattr(mw_row, 'Time_Seconds', 0)
            tc_matches['Time_Diff'] = abs(tc_matches['Time_Seconds'] - mw_time_seconds)
            
            # Filter for time tolerance AND exclude already matched TC rows
            close_time_matches = tc_matches[
                (tc_matches['Time_Diff'] <= time_tolerance) & 
                (~tc_matches.index.isin(already_matched_tc_indices))  
            ].copy()
            
            if close_time_matches.empty:
                continue  # Try next TC mapping
            
            # SUCCESS: Found a match with this TC mapping
            close_time_matches = close_time_matches.sort_values('Time_Diff')
            best_match = close_time_matches.iloc[0]
            match_found = True
            successful_tc_theme = tc_theme_norm
            
            # Track successful mapping attempts
            if match_attempt not in filter_stats["mapping_attempts_breakdown"]:
                filter_stats["mapping_attempts_breakdown"][match_attempt] = 0
            filter_stats["mapping_attempts_breakdown"][match_attempt] += 1
            
            # Log successful fallback matches
            if match_attempt > 1:
                filter_stats["successful_fallback_matches"] += 1
                if i <= 10:  # Only log first 10 for brevity
                    print(f"\n   🎯 Fallback Success #{filter_stats['successful_fallback_matches']}: "
                        f"{mw_row.Advt_Theme} → TC mapping #{match_attempt} ({tc_theme_norm})")
            
            break  # Stop trying other TC mappings
        
        # Handle match results
        if not match_found:
            if match_attempt == 0:
                filter_stats["no_matches_after_theme_filter"] += 1
            else:
                filter_stats["no_matches_after_time_filter"] += 1
            continue
        
        # Track this TC index as already matched
        already_matched_tc_indices.add(best_match.name)
        matched_tc_indices.append(best_match.name)
        
        # Build unique key for matched TC row
        tc_key = (
            str(best_match.get('Advt_Theme', '')).strip().lower(),
            str(best_match.get('Program', '')).strip().lower(),
            str(best_match.get('Advt_time', '')).strip().lower(),
            str(best_match.get('Dd', '')).strip().lower(),
            str(best_match.get('Mn', '')).strip().lower(),
            str(best_match.get('Yr', '')).strip().lower()
        )
        
        # Check for program mismatch
        program_lmrb = mw_row.Program
        program_tc = best_match['Program']
        if str(program_lmrb).strip().lower() != str(program_tc).strip().lower():
            mismatch_row = {
                'Media_Watch_Theme': mw_row.Advt_Theme,
                'TC_Theme': best_match['Advt_Theme'],
                'Date': f"{mw_row.Dd}/{mw_row.Mn}/{mw_row.Yr}",
                'Program_LMRB': program_lmrb,
                'Program_TC': program_tc,
                'Media_Watch_Time': mw_row.Advt_time,
                'TC_Time': best_match['Advt_time'],
                'Time_Difference_Seconds': best_match['Time_Diff'],
                'Channel': getattr(mw_row, 'Channel', ''),
                'Match_Status': 'Time Matched / Program Mismatched',
                'TC_Mapping_Used': match_attempt,  # Track which mapping succeeded
                'Successful_TC_Theme': successful_tc_theme
            }
            program_mismatches.append(mismatch_row)
            filter_stats["program_mismatched_but_time_matched"] += 1
            
        # Always add to matched results
        result_row = {}
        for col in media_watch_df.columns:
            result_row[col] = getattr(mw_row, col)
        
        # Add metadata about which mapping was used
        result_row['TC_Mapping_Attempt_Used'] = match_attempt
        result_row['Successful_TC_Theme'] = successful_tc_theme
        result_row['Total_TC_Options_Available'] = len(tc_mappings_list)
        
        matched_results.append(result_row)
        matched_mw_indices.append(mw_row.Index)
        filter_stats["matches_found"] += 1

    # Print ENHANCED final stats
    print(f"\n\n" + "="*80)
    print("🎉 ENHANCED MATCHING RESULTS")
    print("="*80)
    print(f"📊 Basic Stats:")
    print(f"   Total Media Watch rows processed: {total_rows}")
    print(f"   ✅ Matches found: {filter_stats['matches_found']}")
    print(f"   ❌ No theme mapping: {filter_stats['no_normalized_mapping']}")
    print(f"   ❌ Failed theme filter: {filter_stats['no_matches_after_theme_filter']}")
    print(f"   ❌ Failed date filter: {filter_stats['no_matches_after_date_filter']}")
    print(f"   ❌ Failed time filter: {filter_stats['no_matches_after_time_filter']}")
    print(f"   ⚠️ Program mismatches: {filter_stats['program_mismatched_but_time_matched']}")

    print(f"\n🔀 Enhanced Multiple Mapping Stats:")
    print(f"   Records with multiple TC options: {filter_stats['multiple_mapping_attempts']}")
    print(f"   🎯 Successful fallback matches: {filter_stats['successful_fallback_matches']}")

    if filter_stats["mapping_attempts_breakdown"]:
        print(f"   📈 Mapping attempt breakdown:")
        for attempt_num, count in sorted(filter_stats["mapping_attempts_breakdown"].items()):
            print(f"      Attempt #{attempt_num}: {count} matches")

    # Calculate success rates
    if filter_stats['multiple_mapping_attempts'] > 0:
        fallback_success_rate = (filter_stats['successful_fallback_matches'] / 
                            filter_stats['multiple_mapping_attempts'] * 100)
        print(f"   📊 Fallback success rate: {fallback_success_rate:.1f}%")

    matched_df = pd.DataFrame(matched_results) if matched_results else pd.DataFrame()
    program_mismatched_df = pd.DataFrame(program_mismatches) if program_mismatches else pd.DataFrame()

    print(f"\n🏁 Final Results:")
    print(f"   ✅ Total matches: {len(matched_df)}")
    print(f"   ⚠️ Program mismatches: {len(program_mismatched_df)}")
    print("="*80)

    return matched_df, matched_mw_indices, program_mismatched_df, matched_tc_indices