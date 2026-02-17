/**
 * Sirasa TV PDF Converter
 * Converts Sirasa TV TC PDF to structured data
 */
const sirasaConverter = {
    channelName: 'Sirasa TV',

    /**
     * Convert PDF file to data array
     */
    async convert(file) {
        try {
            const arrayBuffer = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

            const allData = [];

            // Process each page
            for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
                const page = await pdf.getPage(pageNum);
                const textContent = await page.getTextContent();

                // Extract text from page
                const pageText = textContent.items.map(item => item.str).join(' ');

                // Parse text into records
                const pageData = this.parsePage(pageText);
                allData.push(...pageData);
            }

            // Remove duplicates
            const uniqueData = this.removeDuplicates(allData);

            // Sort by date and time
            return this.sortData(uniqueData);

        } catch (error) {
            console.error('Error converting Sirasa TV PDF:', error);
            throw new Error(`PDF conversion failed: ${error.message}`);
        }
    },

    /**
     * Parse page text into records
     */
    parsePage(text) {
        const records = [];
        const lines = text.split('\n');

        lines.forEach(line => {
            // Pattern 1: Date | Program | Time | Theme | Duration
            // Example: 01/01/2024 News 10:30:15:00 Coca Cola 30
            const pattern1 = /(\d{2}\/\d{2}\/\d{4})\s+(.+?)\s+(\d{2}:\d{2}:\d{2}:\d{2})\s+(.+?)\s+(\d+)\s*$/;
            const match1 = line.match(pattern1);

            if (match1) {
                const [_, date, program, time, theme, duration] = match1;
                records.push({
                    Date: date,
                    Program: program.trim(),
                    Time: this.convertTime(time),
                    Theme: theme.trim(),
                    Duration: parseInt(duration)
                });
                return;
            }

            // Pattern 2: Without program name
            // Example: 01/01/2024 10:30:15:00 Coca Cola 30
            const pattern2 = /(\d{2}\/\d{2}\/\d{4})\s+(\d{2}:\d{2}:\d{2}:\d{2})\s+(.+?)\s+(\d+)\s*$/;
            const match2 = line.match(pattern2);

            if (match2) {
                const [_, date, time, theme, duration] = match2;
                records.push({
                    Date: date,
                    Program: 'Unknown',
                    Time: this.convertTime(time),
                    Theme: theme.trim(),
                    Duration: parseInt(duration)
                });
                return;
            }

            // Pattern 3: More flexible format
            const dateMatch = line.match(/(\d{2}\/\d{2}\/\d{4})/);
            const timeMatch = line.match(/(\d{2}:\d{2}:\d{2}(?::\d{2})?)/);
            const durationMatch = line.match(/(\d+)\s*(?:sec|seconds|s)?\s*$/);

            if (dateMatch && timeMatch && durationMatch) {
                const date = dateMatch[1];
                const time = timeMatch[1];
                const duration = durationMatch[1];

                // Extract theme (between time and duration)
                const timeEnd = line.indexOf(time) + time.length;
                const durationStart = line.lastIndexOf(duration);
                const theme = line.substring(timeEnd, durationStart).trim();

                // Extract program (between date and time)
                const dateEnd = line.indexOf(date) + date.length;
                const timeStart = line.indexOf(time);
                const program = line.substring(dateEnd, timeStart).trim();

                if (theme) {
                    records.push({
                        Date: date,
                        Program: program || 'Unknown',
                        Time: this.convertTime(time),
                        Theme: theme,
                        Duration: parseInt(duration)
                    });
                }
            }
        });

        return records;
    },

    /**
     * Convert time format from HH:MM:SS:FF to HH:MM:SS
     */
    convertTime(timeStr) {
        const parts = timeStr.split(':');
        if (parts.length === 4) {
            // HH:MM:SS:FF -> HH:MM:SS
            return `${parts[0]}:${parts[1]}:${parts[2]}`;
        }
        return timeStr;
    },

    /**
     * Remove duplicate records
     */
    removeDuplicates(data) {
        const seen = new Set();
        return data.filter(record => {
            const key = `${record.Date}_${record.Time}_${record.Theme}_${record.Duration}`;
            if (seen.has(key)) {
                return false;
            }
            seen.add(key);
            return true;
        });
    },

    /**
     * Sort data by date and time
     */
    sortData(data) {
        return data.sort((a, b) => {
            // Compare dates
            const dateA = this.parseDate(a.Date);
            const dateB = this.parseDate(b.Date);

            if (dateA < dateB) return -1;
            if (dateA > dateB) return 1;

            // If dates are equal, compare times
            const timeA = this.timeToSeconds(a.Time);
            const timeB = this.timeToSeconds(b.Time);

            return timeA - timeB;
        });
    },

    /**
     * Parse date DD/MM/YYYY
     */
    parseDate(dateStr) {
        const [day, month, year] = dateStr.split('/').map(p => parseInt(p));
        return new Date(year, month - 1, day);
    },

    /**
     * Convert time to seconds
     */
    timeToSeconds(timeStr) {
        const [hours, minutes, seconds] = timeStr.split(':').map(p => parseInt(p));
        return hours * 3600 + minutes * 60 + seconds;
    }
};

// Export for use in app
window.sirasaConverter = sirasaConverter;
