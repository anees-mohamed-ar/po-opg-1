/**
 * Global Configuration for Tally Integration
 * You can edit the Company Name, Tally IP, and Port here directly,
 * or configure them via environment variables.
 */

const config = {
    // Tally Connection Settings
    TALLY_HOST: process.env.TALLY_HOST || 'localhost',
    TALLY_PORT: process.env.TALLY_PORT || 9000,

    // Tally Company Name (must match exactly as displayed in Tally)
    COMPANY_NAME: process.env.TALLY_COMPANY_NAME || 'Opg Legacy Data 22-23',

    // Helper getter for full Tally URL
    get TALLY_URL() {
        if (process.env.TALLY_URL) {
            return process.env.TALLY_URL;
        }
        return `http://${this.TALLY_HOST}:${this.TALLY_PORT}`;
    },

    // Backend Server Port
    SERVER_PORT: process.env.PORT || 5001,
};

module.exports = config;
