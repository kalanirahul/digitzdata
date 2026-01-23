/**
 * DD Consulting - Dynamic Data Loader
 *
 * This module handles loading data from Google Sheets for:
 * - Team Members
 * - Webinars & Events
 * - Training Programs
 * - Trainers
 * - Job Openings
 *
 * SETUP INSTRUCTIONS:
 * 1. Create a Google Sheet with multiple tabs (Team, Webinars, Training, Trainers, Jobs)
 * 2. Publish the sheet: File > Share > Publish to web > Entire Document > CSV
 * 3. Copy the Sheet ID from the URL
 * 4. Update the SHEET_ID below
 *
 * GOOGLE DRIVE IMAGE URLS:
 * To use Google Drive images, convert share links:
 * Original: https://drive.google.com/file/d/FILE_ID/view?usp=sharing
 * Convert to: https://drive.google.com/uc?export=view&id=FILE_ID
 */

const DataLoader = (function() {

  // ============================================
  // CONFIGURATION - UPDATE THESE VALUES
  // ============================================

  // Your Google Sheet ID (from the URL)
  const SHEET_ID = '1M0Pjn2PqD3977Wg0ZSk6iDu3wfYdL8zZAKIIPbb6R58';

  // Google Sheets API Key (optional - for public sheets you can use CSV export)
  const API_KEY = ''; // Leave empty to use CSV export method

  // Sheet tab names (must match exactly)
  const SHEETS = {
    team: 'Team',
    webinars: 'Webinars',
    training: 'Training',
    trainers: 'Trainers',
    jobs: 'Jobs'
  };

  // Cache duration in milliseconds (5 minutes)
  const CACHE_DURATION = 5 * 60 * 1000;

  // Use CORS proxy for local development (set to true when testing locally)
  const USE_CORS_PROXY = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
  const CORS_PROXY = 'https://corsproxy.io/?';

  // ============================================
  // INTERNAL CACHE
  // ============================================

  const cache = {};

  // ============================================
  // UTILITY FUNCTIONS
  // ============================================

  /**
   * Convert Google Drive share link to direct image URL
   */
  function convertGoogleDriveUrl(url) {
    if (!url) return null;

    // Already a direct URL
    if (url.includes('uc?export=view')) {
      return url;
    }

    // Extract file ID from various Google Drive URL formats
    const patterns = [
      /\/file\/d\/([a-zA-Z0-9_-]+)/,
      /id=([a-zA-Z0-9_-]+)/,
      /\/d\/([a-zA-Z0-9_-]+)/
    ];

    for (const pattern of patterns) {
      const match = url.match(pattern);
      if (match && match[1]) {
        return `https://drive.google.com/uc?export=view&id=${match[1]}`;
      }
    }

    // Return original URL if not a Google Drive link
    return url;
  }

  /**
   * Parse CSV string into array of objects
   */
  function parseCSV(csvText) {
    const lines = csvText.split('\n');
    if (lines.length < 2) return [];

    // Parse header row
    const headers = parseCSVLine(lines[0]);

    // Parse data rows
    const data = [];
    for (let i = 1; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line) continue;

      const values = parseCSVLine(line);
      const obj = {};

      headers.forEach((header, index) => {
        if (!header || !header.trim()) return; // Skip empty headers
        const key = toCamelCase(header.trim());
        if (!key) return; // Skip if key is empty
        let value = values[index] || '';

        // Convert image URLs
        if (key.toLowerCase().includes('photo') || key.toLowerCase().includes('image')) {
          value = convertGoogleDriveUrl(value);
        }

        obj[key] = value;
      });

      data.push(obj);
    }

    return data;
  }

  /**
   * Parse a single CSV line, handling quoted values
   */
  function parseCSVLine(line) {
    const values = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
      const char = line[i];

      if (char === '"') {
        inQuotes = !inQuotes;
      } else if (char === ',' && !inQuotes) {
        values.push(current.trim());
        current = '';
      } else {
        current += char;
      }
    }

    values.push(current.trim());
    return values;
  }

  /**
   * Convert header to camelCase
   */
  function toCamelCase(str) {
    if (!str || typeof str !== 'string') return '';
    return str
      .toLowerCase()
      .replace(/[^a-zA-Z0-9]+(.)/g, (_, chr) => chr.toUpperCase())
      .replace(/^./, str[0].toLowerCase());
  }

  /**
   * Check if cache is valid
   */
  function isCacheValid(key) {
    if (!cache[key]) return false;
    return (Date.now() - cache[key].timestamp) < CACHE_DURATION;
  }

  /**
   * Fetch data from Google Sheets
   */
  async function fetchSheetData(sheetName) {
    const cacheKey = `sheet_${sheetName}`;

    // Return cached data if valid
    if (isCacheValid(cacheKey)) {
      return cache[cacheKey].data;
    }

    try {
      // Use CSV export method (works for public sheets)
      // Add cache-busting parameter to force fresh data from Google
      const cacheBuster = Date.now();
      const baseUrl = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}&_=${cacheBuster}`;

      // Use CORS proxy for local development
      const url = USE_CORS_PROXY ? CORS_PROXY + encodeURIComponent(baseUrl) : baseUrl;

      const response = await fetch(url);

      if (!response.ok) {
        throw new Error(`Failed to fetch ${sheetName}: ${response.status}`);
      }

      const csvText = await response.text();
      const data = parseCSV(csvText);

      // Cache the data
      cache[cacheKey] = {
        data,
        timestamp: Date.now()
      };

      return data;

    } catch (error) {
      // Return cached data even if expired, as fallback
      if (cache[cacheKey]) {
        return cache[cacheKey].data;
      }

      // Return demo data for development
      return getDemoData(sheetName);
    }
  }

  // ============================================
  // DEMO DATA (for development/testing)
  // ============================================

  function getDemoData(sheetName) {
    const demoData = {
      [SHEETS.team]: [
        {
          name: 'Husain Feroz Ali',
          role: 'CEO & Founder',
          department: 'exec',
          photo: '',
          linkedin: '',
          bio: 'Fellow of the Society of Actuaries (FSA), USA, with over 20 years of experience in the actuarial field.'
        },
        {
          name: 'Piyush Goel',
          role: 'Actuarial Director',
          department: 'exec',
          photo: '',
          linkedin: '',
          bio: ''
        },
        {
          name: 'Rameez Ali',
          role: 'Associate Director',
          department: 'exec',
          photo: '',
          linkedin: '',
          bio: ''
        }
      ],
      [SHEETS.webinars]: [
        {
          title: 'Demo Webinar (Google Sheets not connected)',
          date: '2026-01-01',
          time: '2:00 PM GMT',
          speakers: 'Demo Speaker (Test Role)',
          description: 'This is demo data. If you see this, Google Sheets is not properly connected. Check sharing settings.',
          status: 'upcoming',
          registerLink: '#',
          recordingLink: ''
        }
      ],
      [SHEETS.training]: [
        {
          title: 'AI for Business Professionals',
          category: 'ai',
          description: 'Learn how to leverage artificial intelligence to transform business operations, improve decision-making, and drive innovation.',
          duration: '3 Days',
          format: 'In-person & Online',
          price: 'Contact Us',
          priceNote: '',
          image: '',
          badge: 'new',
          featured: true,
          nextSession: 'March 2026'
        },
        {
          title: 'Power BI Masterclass',
          category: 'powerbi',
          description: 'Master Microsoft Power BI from fundamentals to advanced visualizations. Create compelling dashboards and reports.',
          duration: '2 Days',
          format: 'Online',
          price: 'Contact Us',
          priceNote: '',
          image: '',
          badge: 'popular',
          featured: false,
          nextSession: 'Flexible'
        },
        {
          title: 'IFRS 17 Implementation',
          category: 'ifrs17',
          description: 'Comprehensive training on IFRS 17 insurance contracts standard implementation, reporting, and compliance.',
          duration: '5 Days',
          format: 'In-person',
          price: 'Contact Us',
          priceNote: '',
          image: '',
          badge: '',
          featured: false,
          nextSession: 'April 2026'
        },
        {
          title: 'Actuarial Foundations',
          category: 'actuarial',
          description: 'Build a strong foundation in actuarial science covering life, health, and general insurance principles.',
          duration: '4 Days',
          format: 'In-person & Online',
          price: 'Contact Us',
          priceNote: '',
          image: '',
          badge: '',
          featured: false,
          nextSession: 'Ongoing'
        },
        {
          title: 'Risk Management Essentials',
          category: 'risk',
          description: 'Learn enterprise risk management frameworks, risk assessment methodologies, and regulatory compliance.',
          duration: '3 Days',
          format: 'Online',
          price: 'Contact Us',
          priceNote: '',
          image: '',
          badge: '',
          featured: false,
          nextSession: 'Flexible'
        },
        {
          title: 'Leadership Development',
          category: 'leadership',
          description: 'Develop essential leadership skills for the modern workplace including communication, strategy, and team management.',
          duration: '2 Days',
          format: 'In-person',
          price: 'Contact Us',
          priceNote: '',
          image: '',
          badge: '',
          featured: false,
          nextSession: 'Quarterly'
        }
      ],
      [SHEETS.trainers]: [
        {
          name: 'Expert Trainer',
          title: 'AI Training Specialist',
          photo: '',
          bio: 'Experienced AI practitioner with expertise in machine learning, natural language processing, and business applications of AI.'
        }
      ],
      [SHEETS.jobs]: [
        {
          title: 'Senior Actuarial Analyst',
          department: 'Actuarial',
          location: 'Dubai, UAE',
          type: 'Full-time',
          description: 'Join our actuarial team to work on cutting-edge insurance projects across the GCC region.',
          requirements: 'FSA/ASA qualification or near-qualified, 3+ years experience, strong Excel and Python skills',
          applyLink: ''
        },
        {
          title: 'Data Scientist',
          department: 'Technology',
          location: 'Dubai, UAE',
          type: 'Full-time',
          description: 'Lead data science initiatives and develop predictive models for insurance applications.',
          requirements: 'Masters in Data Science or related field, 2+ years experience, Python, R, ML frameworks',
          applyLink: ''
        }
      ]
    };

    return demoData[sheetName] || [];
  }

  // ============================================
  // PUBLIC API
  // ============================================

  return {
    /**
     * Load team members from Google Sheets
     */
    async loadTeam() {
      return await fetchSheetData(SHEETS.team);
    },

    /**
     * Load webinars/events from Google Sheets
     */
    async loadWebinars() {
      return await fetchSheetData(SHEETS.webinars);
    },

    /**
     * Load training programs from Google Sheets
     */
    async loadTrainingPrograms() {
      return await fetchSheetData(SHEETS.training);
    },

    /**
     * Load trainers from Google Sheets
     */
    async loadTrainers() {
      return await fetchSheetData(SHEETS.trainers);
    },

    /**
     * Load job openings from Google Sheets
     */
    async loadJobs() {
      return await fetchSheetData(SHEETS.jobs);
    },

    /**
     * Convert Google Drive URL to direct image URL
     */
    convertImageUrl: convertGoogleDriveUrl,

    /**
     * Clear cache (useful for forcing refresh)
     */
    clearCache() {
      Object.keys(cache).forEach(key => delete cache[key]);
    },

    /**
     * Check if the loader is configured
     */
    isConfigured() {
      return SHEET_ID !== 'YOUR_GOOGLE_SHEET_ID_HERE';
    }
  };

})();

// Export for use in other scripts
if (typeof module !== 'undefined' && module.exports) {
  module.exports = DataLoader;
}
