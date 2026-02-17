/**
 * Derana TV PDF Converter
 * Template - Customize the parsing logic for Derana TV TC PDFs
 */
const deranaConverter = {
    channelName: 'Derana TV',

    async convert(file) {
        try {
            const arrayBuffer = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
            const allData = [];

            for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
                const page = await pdf.getPage(pageNum);
                const textContent = await page.getTextContent();
                const pageText = textContent.items.map(item => item.str).join('\n');
                const pageData = this.parsePage(pageText);
                allData.push(...pageData);
            }

            const uniqueData = this.removeDuplicates(allData);
            return this.sortData(uniqueData);
        } catch (error) {
            console.error('Error converting Derana TV PDF:', error);
            throw new Error(`PDF conversion failed: ${error.message}`);
        }
    },

    /**
     * TODO: Customize based on Derana TV PDF format
     */
    parsePage(text) {
        const records = [];
        const lines = text.split('\n');

        lines.forEach(line => {
            // TODO: Add Derana TV specific patterns
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
        });

        return records;
    },

    removeDuplicates(data) {
        const seen = new Set();
        return data.filter(record => {
            const key = `${record.Date}_${record.Time}_${record.Theme}_${record.Duration}`;
            if (seen.has(key)) return false;
            seen.add(key);
            return true;
        });
    },

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

    parseDate(dateStr) {
        const [day, month, year] = dateStr.split('/').map(p => parseInt(p));
        return new Date(year, month - 1, day);
    },

    timeToSeconds(timeStr) {
        const [hours, minutes, seconds] = timeStr.split(':').map(p => parseInt(p) || 0);
        return hours * 3600 + minutes * 60 + (seconds || 0);
    }
};

window.deranaConverter = deranaConverter;
