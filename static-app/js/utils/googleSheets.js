/**
 * Google Sheets API integration
 * Client-side data loading from Google Sheets
 */
const googleSheets = {
    /**
     * Load data from Google Sheet
     */
    async loadData(apiKey, sheetId, startDate, endDate) {
        try {
            const range = 'Sheet1!A:Z'; // Adjust range as needed
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${range}?key=${apiKey}`;

            const response = await fetch(url);

            if (!response.ok) {
                throw new Error('Failed to fetch data from Google Sheets. Check API key and Sheet ID.');
            }

            const result = await response.json();
            const rows = result.values;

            if (!rows || rows.length === 0) {
                throw new Error('No data found in sheet');
            }

            // Convert to array of objects
            const headers = rows[0];
            const data = [];

            for (let i = 1; i < rows.length; i++) {
                const row = {};
                headers.forEach((header, index) => {
                    row[header] = rows[i][index] || '';
                });
                data.push(row);
            }

            // Filter by date range if provided
            let filteredData = data;
            if (startDate || endDate) {
                filteredData = this.filterByDateRange(data, startDate, endDate);
            }

            // Standardize data
            return this.standardizeData(filteredData);

        } catch (error) {
            console.error('Error loading from Google Sheets:', error);
            throw error;
        }
    },

    /**
     * Filter data by date range
     */
    filterByDateRange(data, startDate, endDate) {
        return data.filter(row => {
            if (!row.Date) return true;

            const rowDate = this.parseDate(row.Date);
            const start = startDate ? new Date(startDate) : null;
            const end = endDate ? new Date(endDate) : null;

            if (start && rowDate < start) return false;
            if (end && rowDate > end) return false;

            return true;
        });
    },

    /**
     * Parse date from various formats
     */
    parseDate(dateStr) {
        // Try multiple formats
        const formats = [
            // DD/MM/YYYY
            /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/,
            // MM/DD/YYYY
            /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/,
            // YYYY-MM-DD
            /^(\d{4})-(\d{1,2})-(\d{1,2})$/
        ];

        for (const format of formats) {
            const match = dateStr.match(format);
            if (match) {
                // Assume DD/MM/YYYY for slash format
                if (dateStr.includes('/')) {
                    const [_, day, month, year] = match;
                    return new Date(year, month - 1, day);
                } else {
                    const [_, year, month, day] = match;
                    return new Date(year, month - 1, day);
                }
            }
        }

        // Fallback to Date constructor
        return new Date(dateStr);
    },

    /**
     * Standardize column names and data
     */
    standardizeData(data) {
        return data.map(row => {
            const standardized = {};

            // Standardize common column names
            const columnMap = {
                'date': 'Date',
                'channel': 'Channel',
                'product': 'Product',
                'theme': 'Theme',
                'program': 'Program',
                'programme': 'Program',
                'time': 'Time',
                'duration': 'Duration',
                'advertiser': 'Advertiser'
            };

            Object.keys(row).forEach(key => {
                const lowerKey = key.toLowerCase().trim();
                const standardKey = columnMap[lowerKey] || key;
                standardized[standardKey] = row[key];
            });

            // Ensure Duration is a number
            if (standardized.Duration) {
                standardized.Duration = parseInt(standardized.Duration) || 0;
            }

            return standardized;
        });
    }
};
