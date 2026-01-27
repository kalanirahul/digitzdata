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
    jobs: 'Jobs',
    practices: 'Practices',
    industries: 'Industries'
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

  // Auto-clear cache if URL has ?refresh=1 or ?nocache=1 parameter
  (function checkForceRefresh() {
    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.has('refresh') || urlParams.has('nocache')) {
      Object.keys(cache).forEach(key => delete cache[key]);
      // Remove the parameter from URL without page reload (clean URL)
      if (window.history && window.history.replaceState) {
        urlParams.delete('refresh');
        urlParams.delete('nocache');
        const newUrl = urlParams.toString()
          ? `${window.location.pathname}?${urlParams.toString()}`
          : window.location.pathname;
        window.history.replaceState({}, document.title, newUrl);
      }
    }
  })();

  // Also check sessionStorage for forced refresh flag (set by admin)
  (function checkSessionRefresh() {
    if (sessionStorage.getItem('dd_force_refresh')) {
      Object.keys(cache).forEach(key => delete cache[key]);
      sessionStorage.removeItem('dd_force_refresh');
    }
  })();

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
      ],
      [SHEETS.practices]: [
        {
          name: 'Actuarial Services',
          tagline: 'Risk quantified.',
          homedescription: 'Statutory valuations, pricing, reserving, risk modeling, ORSA, and pension services for life, health, and general insurance.',
          description: 'Our actuarial practice provides comprehensive solutions for life, health, and general insurance companies. We combine deep technical expertise with practical business insight to help clients navigate complex regulatory requirements, optimize pricing strategies, and build robust risk management frameworks.',
          services: 'Appointed Actuary Services|Peer Review Actuary|Pricing & Product Development|Reserving & IBNR Analysis|Risk Modeling|Reinsurance Optimization|M&A Due Diligence|ORSA & Risk Management Services|Actuarial Experts for Auditors|Pension & Gratuity Services',
          icon: 'chart'
        },
        {
          name: 'IFRS 17 Implementation & Support',
          tagline: 'Compliance delivered.',
          homedescription: 'End-to-end IFRS 17 implementation, managed services, training, and actuarial support for insurers.',
          description: 'DD Consulting provides end-to-end IFRS 17 implementation and support services, helping insurers meet regulatory requirements with confidence and efficiency. Our services cover the full IFRS 17 lifecycle, including actuarial assumptions and methodologies, data and input preparation, calculations, results analysis, disclosure preparation, and ongoing regulatory and audit support.',
          services: 'IFRS 17 Managed Services|IFRS 17 Training & Knowledge Transfer|IFRS 17 Resource Outsourcing|Actuarial Expert Services for Auditors|Third-Party IFRS 17 Platforms|In-House IFRS 17 Solutions',
          icon: 'clipboard'
        },
        {
          name: 'Accounting & Finance',
          tagline: 'Clarity in complexity.',
          homedescription: 'IFRS 9 consulting, audit support, financial due diligence, ICOFR, and valuation services.',
          description: 'We provide expert accounting, audit support, and financial advisory services tailored for complex regulatory environments. Our team helps organizations maintain financial integrity while navigating evolving standards and stakeholder expectations.',
          services: 'IFRS 9 Consulting|Audit Support|Financial Due Diligence|ICOFR (Internal Controls Over Financial Reporting)|Financial Reporting|Unlisted Equity Valuation Services',
          icon: 'table'
        },
        {
          name: 'ESG & Sustainability',
          tagline: 'Purpose meets performance.',
          homedescription: 'ESG strategy, climate risk assessment, sustainability reporting, and carbon analysis.',
          description: 'We help organizations measure, report, and improve their environmental, social, and governance impact. Our approach integrates ESG considerations into core business strategy, creating value while addressing stakeholder expectations.',
          services: 'ESG Strategy Development|Climate Risk Assessment|Sustainability Reporting|Carbon Footprint Analysis|ESG Due Diligence|Stakeholder Engagement',
          icon: 'globe'
        },
        {
          name: 'E-Invoicing',
          tagline: 'Compliance automated.',
          homedescription: 'E-invoicing implementation, system integration, and digital invoicing transformation.',
          description: 'Navigate the rapidly evolving landscape of electronic invoicing mandates across the GCC. We help organizations implement compliant systems and processes that meet regulatory requirements while improving operational efficiency.',
          services: 'E-Invoicing Implementation Services|System Integration|Compliance Monitoring|Process Automation|Training & Support',
          icon: 'document'
        },
        {
          name: 'Technology & Analytics',
          tagline: 'Data-driven decisions.',
          homedescription: 'Data engineering, BI dashboards, AI/ML solutions, automation, and cloud reserving platform.',
          description: 'Our technology practice delivers business intelligence, data analytics, and custom software solutions that transform data into competitive advantage. We build tools that solve real problems and create lasting value.',
          services: 'Business Intelligence|Data Engineering|Custom Development|AI & Machine Learning|System Integration|Analytics Strategy|Dashboard & Email Automation Services|Cloud Based Reserving Platform',
          icon: 'monitor'
        },
        {
          name: 'Training & Development',
          tagline: 'Building capability, driving growth.',
          homedescription: 'Actuarial training, IFRS 17 workshops, and custom corporate development programs.',
          description: 'We deliver tailored training programs and professional development solutions that build organizational capability. Our programs combine technical expertise with practical application, empowering teams to excel in their roles and adapt to evolving industry demands.',
          services: 'Actuarial Training Programs|IFRS 17 Workshops|Technical Skills Development|Leadership Development|Regulatory Compliance Training|Custom Corporate Training',
          icon: 'graduation'
        }
      ],
      [SHEETS.industries]: [
        {
          name: 'Insurance',
          description: 'Life, health, and general insurers across the GCC and beyond trust us for actuarial excellence, regulatory compliance, and strategic transformation. We help carriers navigate evolving markets while maintaining profitability.',
          practices: 'Actuarial|Accounting|ESG|Technology',
          icon: 'shield'
        },
        {
          name: 'Reinsurance',
          description: 'We support reinsurers with treaty pricing, reserving, and strategic portfolio optimization. Our deep understanding of risk transfer mechanisms enables better decision-making across the reinsurance value chain.',
          practices: 'Actuarial|Technology',
          icon: 'shield-alert'
        },
        {
          name: 'Banking & Financial Services',
          description: 'We help banks and financial institutions with risk management, regulatory compliance, and digital transformation. From Basel requirements to ESG integration, we address the full spectrum of financial services challenges.',
          practices: 'Accounting|ESG|E-Invoicing|Technology',
          icon: 'table'
        },
        {
          name: 'Corporate & Conglomerates',
          description: 'Large corporations rely on us for employee benefits consulting, ESG strategy, and financial advisory. We help enterprises manage risk, optimize capital, and build sustainable business practices.',
          practices: 'Actuarial|Accounting|ESG|E-Invoicing',
          icon: 'building'
        },
        {
          name: 'Government & Public Sector',
          description: 'We partner with government entities on policy analysis, pension reform, social insurance programs, and digital government initiatives. Our work helps shape public programs that serve citizens effectively.',
          practices: 'Actuarial|ESG|Technology',
          icon: 'government'
        },
        {
          name: 'Healthcare',
          description: 'Healthcare organizations benefit from our expertise in health economics, provider analytics, and population health management. We help payers and providers optimize care delivery while managing costs.',
          practices: 'Actuarial|Technology',
          icon: 'heartbeat'
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
     * Load practices from Google Sheets
     */
    async loadPractices() {
      return await fetchSheetData(SHEETS.practices);
    },

    /**
     * Load industries from Google Sheets
     */
    async loadIndustries() {
      return await fetchSheetData(SHEETS.industries);
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
     * Force refresh on next page load (sets sessionStorage flag)
     * Useful when updating Google Sheets data
     */
    forceRefreshOnNextLoad() {
      sessionStorage.setItem('dd_force_refresh', '1');
    },

    /**
     * Reload current page with fresh data
     */
    refreshNow() {
      Object.keys(cache).forEach(key => delete cache[key]);
      window.location.reload();
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
