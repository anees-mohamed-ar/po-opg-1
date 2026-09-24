// Global Configuration for Frontend Backend API Connection
// Can be configured via environment variables (REACT_APP_BACKEND_HOST, REACT_APP_BACKEND_PORT, REACT_APP_API_BASE)
// or customized directly in this file.

const BACKEND_HOST = process.env.REACT_APP_BACKEND_HOST || window.location.hostname || 'localhost';
const BACKEND_PORT = process.env.REACT_APP_BACKEND_PORT || 5001;

const config = {
  BACKEND_HOST,
  BACKEND_PORT,
  get API_BASE() {
    if (process.env.REACT_APP_API_BASE) {
      return process.env.REACT_APP_API_BASE;
    }
    return `http://${this.BACKEND_HOST}:${this.BACKEND_PORT}`;
  }
};

export default config;
export const API_BASE = config.API_BASE;
