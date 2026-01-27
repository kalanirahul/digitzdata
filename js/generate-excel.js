const ExcelJS = require('exceljs');
const path = require('path');

async function generateExcel() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'DD Consulting';
  workbook.created = new Date();

  // ============ TEAM SHEET ============
  const teamSheet = workbook.addWorksheet('Team');
  teamSheet.columns = [
    { header: 'name', key: 'name', width: 25 },
    { header: 'role', key: 'role', width: 30 },
    { header: 'department', key: 'department', width: 15 },
    { header: 'photo', key: 'photo', width: 50 },
    { header: 'linkedin', key: 'linkedin', width: 40 },
    { header: 'bio', key: 'bio', width: 80 }
  ];

  const teamData = [
    { name: 'Husain Feroz Ali', role: 'CEO & Founder', department: 'exec', photo: '', linkedin: 'https://linkedin.com/in/husainferoz', bio: 'Fellow of the Society of Actuaries (FSA), USA, with over 20 years of experience in the actuarial field.' },
    { name: 'Piyush Goel', role: 'Actuarial Director', department: 'exec', photo: '', linkedin: '', bio: '' },
    { name: 'Rameez Ali', role: 'Associate Director', department: 'exec', photo: '', linkedin: '', bio: '' },
    { name: 'Mohammad Irfan', role: 'Data Scientist & Actuary', department: 'exec', photo: '', linkedin: '', bio: '' },
    { name: 'Kate Bleakley', role: 'Actuary', department: 'exec', photo: '', linkedin: '', bio: '' },
    { name: 'Hasnain Azhar', role: 'Senior Manager', department: 'general', photo: '', linkedin: '', bio: '' },
    { name: 'Rahul Kalani', role: 'Manager', department: 'general', photo: '', linkedin: '', bio: '' },
    { name: 'Uzair Khurshid', role: 'Senior Actuarial Analyst', department: 'general', photo: '', linkedin: '', bio: '' },
    { name: 'Muhammad Maaz', role: 'Actuarial Analyst', department: 'general', photo: '', linkedin: '', bio: '' },
    { name: 'Hamza Saud', role: 'Actuarial Analyst', department: 'general', photo: '', linkedin: '', bio: '' },
    { name: 'Shadab Haider', role: 'Actuarial Analyst', department: 'general', photo: '', linkedin: '', bio: '' },
    { name: 'Salman Khawja', role: 'Consulting Actuary', department: 'life', photo: '', linkedin: '', bio: '' },
    { name: 'Azim Mamdani', role: 'Deputy Manager', department: 'life', photo: '', linkedin: '', bio: '' },
    { name: 'Aun Haider', role: 'Deputy Manager', department: 'life', photo: '', linkedin: '', bio: '' },
    { name: 'Saad Shakeel', role: 'Actuarial Analyst', department: 'life', photo: '', linkedin: '', bio: '' },
    { name: 'Salman Amir', role: 'Actuarial Analyst', department: 'life', photo: '', linkedin: '', bio: '' },
    { name: 'Mohammed Rizwan', role: 'Actuarial Analyst', department: 'life', photo: '', linkedin: '', bio: '' },
    { name: 'Irfan Baig', role: 'IT Security', department: 'tech', photo: '', linkedin: '', bio: '' },
    { name: 'Muhammed Ahmed', role: 'IT Security', department: 'tech', photo: '', linkedin: '', bio: '' },
    { name: 'Mansoob', role: 'Finance Manager', department: 'finance', photo: '', linkedin: '', bio: '' },
    { name: 'Dilshad Feroz', role: 'Accounting & Tax Consultant', department: 'finance', photo: '', linkedin: '', bio: '' },
    { name: 'Faheem Malik', role: 'Investment Analyst', department: 'finance', photo: '', linkedin: '', bio: '' },
    { name: 'Jayesh Bhana', role: 'Sustainability Consultant', department: 'esg', photo: '', linkedin: '', bio: '' },
    { name: 'Rahim Aziz', role: 'Regional Marketing Manager', department: 'marketing', photo: '', linkedin: '', bio: '' }
  ];
  teamData.forEach(row => teamSheet.addRow(row));
  styleHeader(teamSheet);

  // ============ WEBINARS SHEET ============
  const webinarsSheet = workbook.addWorksheet('Webinars');
  webinarsSheet.columns = [
    { header: 'title', key: 'title', width: 40 },
    { header: 'date', key: 'date', width: 15 },
    { header: 'time', key: 'time', width: 15 },
    { header: 'speakers', key: 'speakers', width: 50 },
    { header: 'description', key: 'description', width: 80 },
    { header: 'status', key: 'status', width: 12 },
    { header: 'registerLink', key: 'registerLink', width: 40 },
    { header: 'recordingLink', key: 'recordingLink', width: 40 }
  ];

  const webinarsData = [
    { title: 'AI in Business: The 2026 Landscape', date: '2026-02-28', time: '2:00 PM GMT', speakers: 'Jayson Smith (AI Strategy Consultant), Sarah Lee (Data Scientist)', description: 'Join us for an in-depth exploration of how artificial intelligence is reshaping business strategy, operations, and decision-making in 2026.', status: 'upcoming', registerLink: 'https://forms.google.com/', recordingLink: '' },
    { title: 'Data Strategy for Insurance', date: '2026-03-15', time: '3:00 PM GMT', speakers: 'Mohammad Irfan (Data Scientist & Actuary)', description: 'Master your organization\'s data strategy with practical frameworks and real-world case studies from the insurance industry.', status: 'upcoming', registerLink: 'https://forms.google.com/', recordingLink: '' },
    { title: 'IFRS 17 Implementation Best Practices', date: '2026-04-10', time: '2:00 PM GMT', speakers: 'Piyush Goel (Actuarial Director), Kate Bleakley (Actuary)', description: 'Deep dive into IFRS 17 implementation challenges and solutions. Covers measurement models, transition approaches, and reporting requirements.', status: 'upcoming', registerLink: 'https://forms.google.com/', recordingLink: '' },
    { title: 'Risk Management in Volatile Markets', date: '2026-05-20', time: '3:00 PM GMT', speakers: 'Husain Feroz Ali (CEO & Founder)', description: 'Understanding and managing risk in today\'s uncertain economic environment. Practical tools and frameworks for insurance professionals.', status: 'upcoming', registerLink: 'https://forms.google.com/', recordingLink: '' },
    { title: 'Power BI for Actuaries', date: '2025-12-15', time: '2:00 PM GMT', speakers: 'Rahul Kalani (Manager)', description: 'Recorded session on building actuarial dashboards with Power BI. Covers data modeling, DAX formulas, and visualization best practices.', status: 'recorded', registerLink: '', recordingLink: 'https://youtube.com/watch?v=example' }
  ];
  webinarsData.forEach(row => webinarsSheet.addRow(row));
  styleHeader(webinarsSheet);

  // ============ TRAINING SHEET ============
  const trainingSheet = workbook.addWorksheet('Training');
  trainingSheet.columns = [
    { header: 'title', key: 'title', width: 35 },
    { header: 'category', key: 'category', width: 12 },
    { header: 'description', key: 'description', width: 80 },
    { header: 'duration', key: 'duration', width: 12 },
    { header: 'format', key: 'format', width: 20 },
    { header: 'price', key: 'price', width: 15 },
    { header: 'priceNote', key: 'priceNote', width: 15 },
    { header: 'image', key: 'image', width: 50 },
    { header: 'badge', key: 'badge', width: 10 },
    { header: 'featured', key: 'featured', width: 10 },
    { header: 'nextSession', key: 'nextSession', width: 20 }
  ];

  const trainingData = [
    { title: 'AI for Business Professionals', category: 'ai', description: 'Comprehensive training on leveraging artificial intelligence to transform business operations, improve decision-making, and drive innovation.', duration: '3 Days', format: 'In-person & Online', price: 'Contact Us', priceNote: 'per participant', image: '', badge: 'new', featured: 'true', nextSession: 'March 2026' },
    { title: 'Power BI Masterclass', category: 'powerbi', description: 'Master Microsoft Power BI from fundamentals to advanced visualizations. Learn data modeling, DAX formulas, and create compelling dashboards.', duration: '2 Days', format: 'Online', price: 'Contact Us', priceNote: 'per participant', image: '', badge: 'popular', featured: 'false', nextSession: 'Flexible' },
    { title: 'Qlik Sense Analytics', category: 'qlik', description: 'Complete training on Qlik Sense for business intelligence and data visualization. Build interactive dashboards and master set analysis.', duration: '2 Days', format: 'Online', price: 'Contact Us', priceNote: 'per participant', image: '', badge: '', featured: 'false', nextSession: 'Flexible' },
    { title: 'IFRS 17 Implementation', category: 'ifrs17', description: 'Comprehensive training on IFRS 17 insurance contracts standard. Covers measurement models, transition approaches, and reporting requirements.', duration: '5 Days', format: 'In-person', price: 'Contact Us', priceNote: 'per participant', image: '', badge: '', featured: 'false', nextSession: 'April 2026' },
    { title: 'Actuarial Foundations', category: 'actuarial', description: 'Build a strong foundation in actuarial science covering life, health, and general insurance principles. Includes reserving, pricing, and compliance.', duration: '4 Days', format: 'In-person & Online', price: 'Contact Us', priceNote: 'per participant', image: '', badge: '', featured: 'false', nextSession: 'Ongoing' },
    { title: 'Advanced Reserving Techniques', category: 'actuarial', description: 'Deep dive into IBNR estimation, chain ladder methods, Bornhuetter-Ferguson, and stochastic reserving approaches.', duration: '3 Days', format: 'In-person', price: 'Contact Us', priceNote: 'per participant', image: '', badge: '', featured: 'false', nextSession: 'May 2026' },
    { title: 'Risk Management Essentials', category: 'risk', description: 'Learn enterprise risk management frameworks, risk assessment methodologies, and regulatory compliance including ORSA and stress testing.', duration: '3 Days', format: 'Online', price: 'Contact Us', priceNote: 'per participant', image: '', badge: '', featured: 'false', nextSession: 'Flexible' },
    { title: 'Leadership Development Program', category: 'leadership', description: 'Develop essential leadership skills including strategic communication, team management, decision-making, and change leadership.', duration: '2 Days', format: 'In-person', price: 'Contact Us', priceNote: 'per participant', image: '', badge: '', featured: 'false', nextSession: 'Quarterly' }
  ];
  trainingData.forEach(row => trainingSheet.addRow(row));
  styleHeader(trainingSheet);

  // ============ TRAINERS SHEET ============
  const trainersSheet = workbook.addWorksheet('Trainers');
  trainersSheet.columns = [
    { header: 'name', key: 'name', width: 25 },
    { header: 'title', key: 'title', width: 35 },
    { header: 'photo', key: 'photo', width: 50 },
    { header: 'bio', key: 'bio', width: 100 }
  ];

  const trainersData = [
    { name: 'Husain Feroz Ali', title: 'CEO & Founder - Lead Trainer', photo: '', bio: 'Fellow of the Society of Actuaries (FSA), USA, with over 20 years of experience. Specializes in actuarial science, risk management, and strategic consulting.' },
    { name: 'Mohammad Irfan', title: 'Data Science & AI Trainer', photo: '', bio: 'Expert in machine learning, data analytics, and AI applications for insurance. Combines actuarial expertise with cutting-edge technology skills.' },
    { name: 'Piyush Goel', title: 'IFRS 17 & Actuarial Trainer', photo: '', bio: 'Actuarial Director with extensive experience in IFRS 17 implementation, reserving, and regulatory compliance across multiple markets.' }
  ];
  trainersData.forEach(row => trainersSheet.addRow(row));
  styleHeader(trainersSheet);

  // ============ JOBS SHEET ============
  const jobsSheet = workbook.addWorksheet('Jobs');
  jobsSheet.columns = [
    { header: 'title', key: 'title', width: 30 },
    { header: 'department', key: 'department', width: 20 },
    { header: 'location', key: 'location', width: 20 },
    { header: 'type', key: 'type', width: 15 },
    { header: 'description', key: 'description', width: 80 },
    { header: 'requirements', key: 'requirements', width: 80 },
    { header: 'applyLink', key: 'applyLink', width: 50 }
  ];

  const jobsData = [
    { title: 'Senior Actuarial Analyst', department: 'Actuarial', location: 'Dubai, UAE', type: 'Full-time', description: 'Join our actuarial team to work on cutting-edge insurance projects across the GCC region.', requirements: 'FSA/ASA qualification or near-qualified, 3+ years experience, strong Excel and Python skills', applyLink: '' },
    { title: 'Data Scientist', department: 'Technology', location: 'Dubai, UAE', type: 'Full-time', description: 'Lead data science initiatives and develop predictive models for insurance applications.', requirements: 'Masters in Data Science or related field, 2+ years experience, Python, R, ML frameworks', applyLink: '' },
    { title: 'IFRS 17 Consultant', department: 'Actuarial', location: 'Riyadh, KSA', type: 'Full-time', description: 'Support IFRS 17 implementation projects for insurance clients across the region.', requirements: 'IFRS 17 implementation experience, actuarial background preferred, excellent communication skills', applyLink: '' }
  ];
  jobsData.forEach(row => jobsSheet.addRow(row));
  styleHeader(jobsSheet);

  // ============ PRACTICES SHEET ============
  const practicesSheet = workbook.addWorksheet('Practices');
  practicesSheet.columns = [
    { header: 'name', key: 'name', width: 35 },
    { header: 'tagline', key: 'tagline', width: 30 },
    { header: 'homeDescription', key: 'homeDescription', width: 80 },
    { header: 'description', key: 'description', width: 100 },
    { header: 'services', key: 'services', width: 120 },
    { header: 'icon', key: 'icon', width: 15 }
  ];

  const practicesData = [
    { name: 'Actuarial Services', tagline: 'Risk quantified.', homeDescription: 'Statutory valuations, pricing, reserving, risk modeling, ORSA, and pension services for life, health, and general insurance.', description: 'Our actuarial practice provides comprehensive solutions for life, health, and general insurance companies. We combine deep technical expertise with practical business insight to help clients navigate complex regulatory requirements, optimize pricing strategies, and build robust risk management frameworks.', services: 'Appointed Actuary Services|Peer Review Actuary|Pricing & Product Development|Reserving & IBNR Analysis|Risk Modeling|Reinsurance Optimization|M&A Due Diligence|ORSA & Risk Management Services|Actuarial Experts for Auditors|Pension & Gratuity Services', icon: 'chart' },
    { name: 'IFRS 17 Implementation & Support', tagline: 'Compliance delivered.', homeDescription: 'End-to-end IFRS 17 implementation, managed services, training, and actuarial support for insurers.', description: 'DD Consulting provides end-to-end IFRS 17 implementation and support services, helping insurers meet regulatory requirements with confidence and efficiency. Our services cover the full IFRS 17 lifecycle, including actuarial assumptions and methodologies, data and input preparation, calculations, results analysis, disclosure preparation, and ongoing regulatory and audit support.', services: 'IFRS 17 Managed Services|IFRS 17 Training & Knowledge Transfer|IFRS 17 Resource Outsourcing|Actuarial Expert Services for Auditors|Third-Party IFRS 17 Platforms|In-House IFRS 17 Solutions', icon: 'clipboard' },
    { name: 'Accounting & Finance', tagline: 'Clarity in complexity.', homeDescription: 'IFRS 9 consulting, audit support, financial due diligence, ICOFR, and valuation services.', description: 'We provide expert accounting, audit support, and financial advisory services tailored for complex regulatory environments. Our team helps organizations maintain financial integrity while navigating evolving standards and stakeholder expectations.', services: 'IFRS 9 Consulting|Audit Support|Financial Due Diligence|ICOFR (Internal Controls Over Financial Reporting)|Financial Reporting|Unlisted Equity Valuation Services', icon: 'table' },
    { name: 'ESG & Sustainability', tagline: 'Purpose meets performance.', homeDescription: 'ESG strategy, climate risk assessment, sustainability reporting, and carbon analysis.', description: 'We help organizations measure, report, and improve their environmental, social, and governance impact. Our approach integrates ESG considerations into core business strategy, creating value while addressing stakeholder expectations.', services: 'ESG Strategy Development|Climate Risk Assessment|Sustainability Reporting|Carbon Footprint Analysis|ESG Due Diligence|Stakeholder Engagement', icon: 'globe' },
    { name: 'E-Invoicing', tagline: 'Compliance automated.', homeDescription: 'E-invoicing implementation, system integration, and digital invoicing transformation.', description: 'Navigate the rapidly evolving landscape of electronic invoicing mandates across the GCC. We help organizations implement compliant systems and processes that meet regulatory requirements while improving operational efficiency.', services: 'E-Invoicing Implementation Services|System Integration|Compliance Monitoring|Process Automation|Training & Support', icon: 'document' },
    { name: 'Technology & Analytics', tagline: 'Data-driven decisions.', homeDescription: 'Data engineering, BI dashboards, AI/ML solutions, automation, and cloud reserving platform.', description: 'Our technology practice delivers business intelligence, data analytics, and custom software solutions that transform data into competitive advantage. We build tools that solve real problems and create lasting value.', services: 'Business Intelligence|Data Engineering|Custom Development|AI & Machine Learning|System Integration|Analytics Strategy|Dashboard & Email Automation Services|Cloud Based Reserving Platform', icon: 'monitor' },
    { name: 'Training & Development', tagline: 'Building capability, driving growth.', homeDescription: 'Actuarial training, IFRS 17 workshops, and custom corporate development programs.', description: 'We deliver tailored training programs and professional development solutions that build organizational capability. Our programs combine technical expertise with practical application, empowering teams to excel in their roles and adapt to evolving industry demands.', services: 'Actuarial Training Programs|IFRS 17 Workshops|Technical Skills Development|Leadership Development|Regulatory Compliance Training|Custom Corporate Training', icon: 'graduation' }
  ];
  practicesData.forEach(row => practicesSheet.addRow(row));
  styleHeader(practicesSheet);

  // ============ INDUSTRIES SHEET ============
  const industriesSheet = workbook.addWorksheet('Industries');
  industriesSheet.columns = [
    { header: 'name', key: 'name', width: 35 },
    { header: 'description', key: 'description', width: 100 },
    { header: 'practices', key: 'practices', width: 50 },
    { header: 'icon', key: 'icon', width: 15 }
  ];

  const industriesData = [
    { name: 'Insurance', description: 'Life, health, and general insurers across the GCC and beyond trust us for actuarial excellence, regulatory compliance, and strategic transformation. We help carriers navigate evolving markets while maintaining profitability.', practices: 'Actuarial|Accounting|ESG|Technology', icon: 'shield' },
    { name: 'Reinsurance', description: 'We support reinsurers with treaty pricing, reserving, and strategic portfolio optimization. Our deep understanding of risk transfer mechanisms enables better decision-making across the reinsurance value chain.', practices: 'Actuarial|Technology', icon: 'shield-alert' },
    { name: 'Banking & Financial Services', description: 'We help banks and financial institutions with risk management, regulatory compliance, and digital transformation. From Basel requirements to ESG integration, we address the full spectrum of financial services challenges.', practices: 'Accounting|ESG|E-Invoicing|Technology', icon: 'table' },
    { name: 'Corporate & Conglomerates', description: 'Large corporations rely on us for employee benefits consulting, ESG strategy, and financial advisory. We help enterprises manage risk, optimize capital, and build sustainable business practices.', practices: 'Actuarial|Accounting|ESG|E-Invoicing', icon: 'building' },
    { name: 'Government & Public Sector', description: 'We partner with government entities on policy analysis, pension reform, social insurance programs, and digital government initiatives. Our work helps shape public programs that serve citizens effectively.', practices: 'Actuarial|ESG|Technology', icon: 'government' },
    { name: 'Healthcare', description: 'Healthcare organizations benefit from our expertise in health economics, provider analytics, and population health management. We help payers and providers optimize care delivery while managing costs.', practices: 'Actuarial|Technology', icon: 'heartbeat' }
  ];
  industriesData.forEach(row => industriesSheet.addRow(row));
  styleHeader(industriesSheet);

  // ============ INSTRUCTIONS SHEET ============
  const instructionsSheet = workbook.addWorksheet('Instructions');
  instructionsSheet.columns = [
    { header: 'Instructions', key: 'text', width: 100 }
  ];

  const instructions = [
    '',
    'DD CONSULTING WEBSITE DATA - SETUP INSTRUCTIONS',
    '================================================',
    '',
    'STEP 1: UPLOAD TO GOOGLE SHEETS',
    '--------------------------------',
    '1. Go to sheets.google.com',
    '2. Click File > Import > Upload this Excel file',
    '3. Choose "Replace spreadsheet" or "Create new spreadsheet"',
    '',
    'STEP 2: PUBLISH YOUR SHEET',
    '--------------------------',
    '1. Click File > Share > Publish to web',
    '2. Select "Entire Document"',
    '3. Click "Publish"',
    '',
    'STEP 3: GET YOUR SHEET ID',
    '-------------------------',
    'Your Google Sheet URL: https://docs.google.com/spreadsheets/d/SHEET_ID_HERE/edit',
    'Copy the SHEET_ID (the long string between /d/ and /edit)',
    '',
    'STEP 4: UPDATE WEBSITE CODE',
    '---------------------------',
    '1. Open js/data-loader.js',
    '2. Find: const SHEET_ID = \'YOUR_GOOGLE_SHEET_ID_HERE\';',
    '3. Replace with your actual Sheet ID',
    '',
    '================================================',
    'COLUMN REFERENCE',
    '================================================',
    '',
    'TEAM - Department Values:',
    '  exec = Leadership',
    '  general = Actuarial - General',
    '  life = Actuarial - Life',
    '  tech = Technology',
    '  finance = Finance',
    '  esg = ESG',
    '  marketing = Marketing',
    '',
    'WEBINARS - Status Values:',
    '  upcoming = Future event (shows Register button)',
    '  recorded = Past event (shows Request Recording button)',
    '',
    'WEBINARS - Register Link:',
    '  1. Create a Google Form for registration',
    '  2. Copy the form URL (e.g., https://forms.google.com/...)',
    '  3. Paste the URL in the registerLink column',
    '  4. When users click Register, they go to your form',
    '  5. If empty, defaults to contact page',
    '',
    'WEBINARS - Recording Link:',
    '  1. For recorded sessions, add the recording URL',
    '  2. If empty, Request Recording links to contact page',
    '',
    'WEBINARS - Multiple Speakers Format:',
    '  Name1 (Role1), Name2 (Role2)',
    '',
    'TRAINING - Category Values:',
    '  ai, powerbi, qlik, ifrs17, actuarial, risk, leadership',
    '',
    'TRAINING - Badge Values:',
    '  new = Shows NEW label',
    '  popular = Shows POPULAR label',
    '  (empty) = No label',
    '',
    'TRAINING - Featured:',
    '  true = Featured program at top',
    '  false = Regular grid display',
    '',
    'JOBS - Column Reference:',
    '  title = Job title (e.g., Senior Actuarial Analyst)',
    '  department = Department name (e.g., Actuarial, Technology)',
    '  location = Work location (e.g., Dubai, UAE)',
    '  type = Employment type (Full-time, Part-time, Contract, Remote)',
    '  description = Brief job description',
    '  requirements = Key requirements (comma-separated)',
    '  applyLink = Application URL or Google Form link (if empty, links to email)',
    '',
    'JOBS - Notes:',
    '  - Jobs section only shows if there are entries in the sheet',
    '  - Delete all rows (except header) to hide the job listings section',
    '  - Apply button links to applyLink URL or defaults to email if empty',
    '',
    'PRACTICES - Column Reference:',
    '  name = Practice name (e.g., Actuarial Services)',
    '  tagline = Short tagline with period (e.g., Risk quantified.) - shown on Practices page',
    '  homeDescription = Medium description for Home page cards (1-2 sentences)',
    '  description = Full practice description paragraph - shown on Practices page',
    '  services = Pipe-separated list (e.g., Service 1|Service 2|Service 3)',
    '  icon = Icon name: chart, clipboard, table, globe, document, monitor, graduation',
    '',
    'PRACTICES - Icon Values:',
    '  chart = Line chart (Actuarial)',
    '  clipboard = Clipboard/checklist (IFRS 17)',
    '  table = Grid/table (Accounting)',
    '  globe = Globe/world (ESG)',
    '  document = Document (E-Invoicing)',
    '  monitor = Computer screen (Technology)',
    '  graduation = Graduation cap (Training)',
    '',
    'INDUSTRIES - Column Reference:',
    '  name = Industry name (e.g., Insurance)',
    '  description = Industry description paragraph',
    '  practices = Pipe-separated practice tags (e.g., Actuarial|Accounting|ESG)',
    '  icon = Icon name: shield, shield-alert, table, building, government, heartbeat',
    '',
    'INDUSTRIES - Icon Values:',
    '  shield = Shield (Insurance)',
    '  shield-alert = Shield with alert (Reinsurance)',
    '  table = Grid/table (Banking)',
    '  building = Building (Corporate)',
    '  government = Government building (Public Sector)',
    '  heartbeat = Heart pulse (Healthcare)',
    '',
    'PHOTOS - Local Images (Recommended):',
    '  1. Add photos to images/team/ folder in website',
    '  2. Use path: images/team/firstname-lastname.jpg',
    '  3. Example: images/team/rahul-kalani.jpg',
    '',
    'PHOTOS - Google Drive (Alternative):',
    '  1. Upload to Google Drive',
    '  2. Share > Anyone with link > Viewer',
    '  3. Use URL: https://drive.google.com/uc?export=view&id=FILE_ID',
    '  Note: Google Drive images may be unreliable'
  ];

  instructions.forEach(text => {
    instructionsSheet.addRow({ text });
  });
  instructionsSheet.getColumn(1).font = { name: 'Consolas', size: 11 };

  // Save the file
  const outputPath = path.join(__dirname, 'DD-Consulting-Website-Data.xlsx');
  await workbook.xlsx.writeFile(outputPath);
  console.log(`Excel file created: ${outputPath}`);
}

function styleHeader(sheet) {
  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF5A7A8A' }  // DD Consulting teal color
  };
  headerRow.alignment = { vertical: 'middle', horizontal: 'left' };
  sheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: sheet.columns.length }
  };
}

generateExcel().catch(console.error);
