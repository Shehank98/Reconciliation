/**
 * LocalStorage utility for saving/loading data
 */
const storage = {
    /**
     * Save data to localStorage
     */
    save(key, data) {
        try {
            const jsonData = JSON.stringify(data);
            localStorage.setItem(`reconciliation_${key}`, jsonData);
            return true;
        } catch (error) {
            console.error('Error saving to localStorage:', error);
            return false;
        }
    },

    /**
     * Load data from localStorage
     */
    load(key) {
        try {
            const jsonData = localStorage.getItem(`reconciliation_${key}`);
            return jsonData ? JSON.parse(jsonData) : null;
        } catch (error) {
            console.error('Error loading from localStorage:', error);
            return null;
        }
    },

    /**
     * Remove item from localStorage
     */
    remove(key) {
        localStorage.removeItem(`reconciliation_${key}`);
    },

    /**
     * Clear all reconciliation data
     */
    clearAll() {
        const keys = Object.keys(localStorage);
        keys.forEach(key => {
            if (key.startsWith('reconciliation_')) {
                localStorage.removeItem(key);
            }
        });
    }
};
