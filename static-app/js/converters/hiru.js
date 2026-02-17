/**
 * Hiru TV PDF Converter
 * Template - Customize the parsing logic for Hiru TV TC PDFs
 */
const hiruConverter = {
    channelName: 'Hiru TV',

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
                const pageText = textContent.items.map(item => item.str).join('\n');

                // Parse text into records
                const pageData = this.parsePage(pageText);
                allData.push(...pageData);
            }

            // Remove duplicates
            const uniqueData = this.removeDuplicates(allData);

            // Sort by date and time
            return this.sortData(uniqueData);

        } catch (error) {
            console.error('Error converting Hiru TV PDF:', error);
            throw new Error(`PDF conversion failed: ${error.message}`);
        }
    },

    /**
     * Parse page text into records
     * TODO: Customize this method based on Hiru TV PDF format
     */
    parsePage(text) {
        const records = [];
        const lines = text.split('\n');

        lines.forEach(line => {
            // TODO: Add your Hiru TV specific parsing patterns here
            // Example patterns to detect:
            // - Date format: DD/MM/YYYY or MM/DD/YYYY or YYYY-MM-DD?
            // - Time format: HH:MM:SS or HH:MM or other?
            // - How are columns separated? Spaces, tabs, pipes?
            // - What is the order of fields?

            // Template pattern (modify based on actual format):
            const pattern = /(\d{2}\/\d{2}\/\d{4})\s+(.+?)\s+(\d{2}:\d{2}:\d{2})\s+(.+?)\s+(\d+)/;
            const match = line.match(pattern);

            if (match) {
                const [_, date, program, time, theme, duration] = match;
                records.push({
                    Date: date,
                    Program: program.trim(),
                    Time: time,
                    Theme: theme.trim(),
                    Duration: parseInt(duration)
                });
            }

            // Add more patterns as needed for different line formats
        });

        return records;
    },

    /**
     * Remove duplicate records
     */
    removeDuplicates(data) {
        const seen = new Set();
        return data.filter(record => {
            const key = `${record.Date}_${record.Time}_${record.Theme}_${record.Duration}`;
            if (seen.has(key)) return false;
            seen.add(key);
            return true;
        });
    },

    /**
     * Sort data by date and time
     */
    sortData(data) {
        return data.sort((a, b) => {
            const dateA = this.parseDate(a.Date);
            const dateB = this.parseDate(b.Date);
            if (dateA < dateB) return -1;
            if (dateA > dateB) return 1;

            const timeA = this.timeToSeconds(a.Time);
            const timeB = this.timeToSeconds(b.Time);
            return timeA - timeB;
        });
    },

    /**
     * Parse date - adjust format based on Hiru TV
     */
    parseDate(dateStr) {
        // TODO: Adjust date parsing based on Hiru TV format
        const [day, month, year] = dateStr.split('/').map(p => parseInt(p));
        return new Date(year, month - 1, day);
    },

    /**
     * Convert time to seconds
     */
    timeToSeconds(timeStr) {
        const [hours, minutes, seconds] = timeStr.split(':').map(p => parseInt(p) || 0);
        return hours * 3600 + minutes * 60 + (seconds || 0);
    }
};

window.hiruConverter = hiruConverter;
