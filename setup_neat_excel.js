const XLSX = require('xlsx');

const EXCEL_FILE = './accounts.xlsx';

const headers = [
  'Email', 'Password', 'First Name', 'Last Name', 'Company Name', 
  'Company Size', 'Phone', 'Job Title', 'Address', 'City', 
  'State', 'Postal Code', 'Country', 'Card Number', 'CVV', 
  'Exp Month', 'Exp Year', 'Status', 'Domain Email'
];

const data = [
  [
    'orion.wren@detroitautomotive.us', 'AutoOrion6789@Detroit', 'Orion', 'Wren', 'Detroit Automotive Design',
    '1 person', '+13135556789', 'Automotive Engineer', '1500 Woodward Avenue', 'Detroit',
    'Michigan', '48226', 'United States', '5198939816602718', '213',
    '03', '30', '', ''
  ],
  [
    'peregrine.quinn@minneapolismedical.us', 'MedicalPeregrine1234@Minneapolis', 'Peregrine', 'Quinn', 'Minneapolis Medical Center',
    '1 person', '+16125551234', 'Cardiologist', '2001 6th Street SE', 'Minneapolis',
    'Minnesota', '55455', 'United States', '5198939816602718', '213',
    '03', '30', '', ''
  ],
  [
    'quill.sterling@stlouisresearch.us', 'BioQuill8901@StLouis', 'Quill', 'Sterling', 'St Louis Research Institute',
    '1 person', '+13145558901', 'Biomedical Researcher', '4444 Forest Park Avenue', 'St Louis',
    'Missouri', '63108', 'United States', '5198939816602718', '213',
    '03', '30', '', ''
  ],
  [
    'raven.black@baltimoremarine.us', 'MarineRaven6789@Baltimore', 'Raven', 'Black', 'Baltimore Marine Institute',
    '1 person', '+14105556789', 'Marine Biologist', '601 East Pratt Street', 'Baltimore',
    'Maryland', '21202', 'United States', '5198939816602718', '213',
    '03', '30', '', ''
  ],
  [
    'sylvan.reed@phoenixsolar.us', 'SolarSylvan1234@Phoenix', 'Sylvan', 'Reed', 'Phoenix Solar Energy',
    '1 person', '+16025551234', 'Solar Engineer', '234 North Central Avenue', 'Phoenix',
    'Arizona', '85004', 'United States', '5198939816602718', '213',
    '03', '30', '', ''
  ],
  [
    'thorne.wilder@nashvillemusic.us', 'MusicThorne8901@Nashville', 'Thorne', 'Wilder', 'Nashville Music Group',
    '1 person', '+16155558901', 'Music Producer', '222 5th Avenue South', 'Nashville',
    'Tennessee', '37203', 'United States', '5198939816602718', '213',
    '03', '30', '', ''
  ]
];

const wb = XLSX.utils.book_new();
const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);

// SET LEBAR KOLOM (WCH = Width Character)
// Ini yang bikin file kelihatan "Rapi" karena kolomnya pas dengan isi
ws['!cols'] = [
  { wch: 35 }, // Email
  { wch: 35 }, // Password
  { wch: 15 }, // First Name
  { wch: 15 }, // Last Name
  { wch: 30 }, // Company Name
  { wch: 15 }, // Company Size
  { wch: 15 }, // Phone
  { wch: 25 }, // Job Title
  { wch: 30 }, // Address
  { wch: 15 }, // City
  { wch: 15 }, // State
  { wch: 12 }, // Postal Code
  { wch: 15 }, // Country
  { wch: 20 }, // Card Number
  { wch: 8 },  // CVV
  { wch: 10 }, // Exp Month
  { wch: 10 }, // Exp Year
  { wch: 12 }, // Status
  { wch: 40 }  // Domain Email
];

XLSX.utils.book_append_sheet(wb, ws, 'Accounts');
XLSX.writeFile(wb, EXCEL_FILE);

console.log('✅ accounts.xlsx has been recreated and it is now neat (with proper column widths)!');
