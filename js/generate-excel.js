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
