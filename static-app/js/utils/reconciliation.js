/**
 * Reconciliation Engine
 * Matches TC data with LMRB data based on theme mappings
 */
const reconciliation = {
    /**
     * Match LMRB and TC data
     */
    match(lmrbData, tcData, mappings, timeTolerance = 30, ignoreDate = false) {
        const matched = [];
        const matchedLMRBIndices = new Set();
        const matchedTCIndices = new Set();

        // Create mapping index
        const mappingIndex = this.createMappingIndex(mappings);

        // Iterate through LMRB data
        lmrbData.forEach((lmrbRow, lmrbIndex) => {
            if (matchedLMRBIndices.has(lmrbIndex)) return;

            const lmrbTheme = this.normalizeTheme(lmrbRow.Theme);
            const lmrbDuration = parseInt(lmrbRow.Duration) || 0;
            const lmrbKey = `${lmrbTheme}_${lmrbDuration}`;

            // Check if there's a mapping for this LMRB theme
            const tcMappings = mappingIndex[lmrbKey];
            if (!tcMappings) return;

            // Try to find matching TC record
            for (const tcMapping of tcMappings) {
                const matchedTC = this.findMatchingTC(
                    lmrbRow,
                    tcData,
                    tcMapping,
                    matchedTCIndices,
                    timeTolerance,
                    ignoreDate
                );

                if (matchedTC) {
                    matched.push({
                        ...lmrbRow,
                        TC_Theme: matchedTC.row.Theme,
                        TC_Time: matchedTC.row.Time,
                        TC_Duration: matchedTC.row.Duration,
                        TC_Program: matchedTC.row.Program,
                        MappingType: tcMapping.type,
                        MatchType: 'AUTO'
                    });

                    matchedLMRBIndices.add(lmrbIndex);
                    matchedTCIndices.add(matchedTC.index);
                    break;
                }
            }
        });

        // Get unmatched records
        const unmatchedLMRB = lmrbData.filter((_, index) => !matchedLMRBIndices.has(index));
        const unmatchedTC = tcData.filter((_, index) => !matchedTCIndices.has(index));

        return {
            matched,
            unmatchedLMRB,
            unmatchedTC
        };
    },

    /**
     * Create mapping index for fast lookup
     */
    createMappingIndex(mappings) {
        const index = {};

        mappings.forEach(mapping => {
            const lmrbKey = `${this.normalizeTheme(mapping.lmrbTheme)}_${mapping.lmrbDuration}`;

            if (!index[lmrbKey]) {
                index[lmrbKey] = [];
            }

            index[lmrbKey].push({
                tcTheme: this.normalizeTheme(mapping.tcTheme),
                tcDuration: mapping.tcDuration,
                type: mapping.type
            });
        });

        return index;
    },

    /**
     * Find matching TC record
     */
    findMatchingTC(lmrbRow, tcData, tcMapping, matchedIndices, timeTolerance, ignoreDate) {
        for (let i = 0; i < tcData.length; i++) {
            if (matchedIndices.has(i)) continue;

            const tcRow = tcData[i];

            // Check theme match
            const tcTheme = this.normalizeTheme(tcRow.Theme);
            if (tcTheme !== tcMapping.tcTheme) continue;

            // Check duration match
            const tcDuration = parseInt(tcRow.Duration) || 0;
            if (tcDuration !== tcMapping.tcDuration) continue;

            // Check date match (if not ignored)
            if (!ignoreDate) {
                if (!this.datesMatch(lmrbRow.Date, tcRow.Date)) continue;
            }

            // Check time match within tolerance
            if (!this.timesMatch(lmrbRow.Time, tcRow.Time, timeTolerance)) continue;

            // Found a match!
            return { row: tcRow, index: i };
        }

        return null;
    },

    /**
     * Normalize theme name
     */
    normalizeTheme(theme) {
        if (!theme) return '';

        return theme
            .toUpperCase()
            .trim()
            .replace(/\s+/g, ' ')
            .replace(/[^\w\s]/g, '');
    },

    /**
     * Check if dates match
     */
    datesMatch(date1, date2) {
        if (!date1 || !date2) return false;

        const d1 = this.parseDate(date1);
        const d2 = this.parseDate(date2);

        return d1.toDateString() === d2.toDateString();
    },

    /**
     * Check if times match within tolerance
     */
    timesMatch(time1, time2, tolerance) {
        if (!time1 || !time2) return false;

        const seconds1 = this.timeToSeconds(time1);
        const seconds2 = this.timeToSeconds(time2);

        const diff = Math.abs(seconds1 - seconds2);

        return diff <= tolerance;
    },

    /**
     * Convert time string to seconds
     */
    timeToSeconds(timeStr) {
        const parts = timeStr.split(':').map(p => parseInt(p) || 0);

        if (parts.length === 3) {
            // HH:MM:SS
            return parts[0] * 3600 + parts[1] * 60 + parts[2];
        } else if (parts.length === 2) {
            // HH:MM
            return parts[0] * 3600 + parts[1] * 60;
        }

        return 0;
    },

    /**
     * Parse date from various formats
     */
    parseDate(dateStr) {
        // Try DD/MM/YYYY format
        const slashMatch = dateStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (slashMatch) {
            const [_, day, month, year] = slashMatch;
            return new Date(year, month - 1, day);
        }

        // Try YYYY-MM-DD format
        const dashMatch = dateStr.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
        if (dashMatch) {
            const [_, year, month, day] = dashMatch;
            return new Date(year, month - 1, day);
        }

        // Fallback
        return new Date(dateStr);
    }
};
