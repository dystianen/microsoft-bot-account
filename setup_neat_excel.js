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
