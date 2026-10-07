/*
 * course.js — course-wide facts (title, units, grading, tools) and the shared
 * datasets reused across weeks. Datasets are generated from a fixed seed so
 * every student sees exactly the same rows, every time.
 */
(function (root) {
  'use strict';

  const C = (root.DIM = root.DIM || {}).Content = root.DIM.Content || { lessons: [], quizzes: [], exams: [] };
  const isNode = typeof module === 'object' && module.exports;
  const F = isNode ? require('../core/formula.js') : root.DIM.Formula;

  C.course = {
    title: 'Intro to Data and Information Management',
    short: 'Data & Info Mgmt',
    grade: '12th grade · one semester · 15 weeks',
    description: 'How organizations collect, store, organize and use data to make decisions — learned hands-on in Excel, then made repeatable with Power Query and Power BI.',
    outcomes: [
      'Explain the difference between data, information and knowledge, and describe how an organization turns data into a decision.',
      'Classify data by type (nominal, ordinal, discrete, continuous; structured vs unstructured) and store it correctly in Excel.',
      'Navigate Excel confidently and build formulas with relative and absolute references.',
      'Choose between CSV and Excel formats, and import, split and export delimited data without losing information.',
      'Assess and fix data quality problems (blanks, duplicates, inconsistent text, wrong types).',
      'Summarize data with conditional functions, lookups and pivot tables, and explain primary and foreign keys.',
      'Pick the right chart for a question and spot a misleading one.',
      'Build a repeatable cleaning process in Power Query and a simple, honest dashboard in Power BI.',
      'Discuss the privacy and ethical responsibilities that come with handling data.',
    ],
    units: [
      { id: 1, title: 'Data foundations and Excel basics', weeks: [1, 2, 3, 4, 5], exam: 'exam-1', color: 'u1' },
      { id: 2, title: 'Analyzing and visualizing data', weeks: [6, 7, 8, 9, 10], exam: 'exam-2', color: 'u2' },
      { id: 3, title: 'Repeatable data work with Power Query and Power BI', weeks: [11, 12, 13, 14, 15], exam: 'exam-3', color: 'u3' },
    ],
    rhythm: [
      ['Monday', '10–15 minute quiz on the previous week (weeks 2–5, 7–10, 12–15). No quiz the Monday after an exam.'],
      ['Tuesday–Thursday', 'Mini-lesson taught inside a live workbook, guided practice, in-class lab.'],
      ['Friday', 'Lab time and catch-up, or the unit exam in weeks 5, 10 and 15.'],
      ['Sunday 11:59 pm', 'Weekly assignment due, submitted as a workbook so the work itself is graded.'],
    ],
    grading: [
      ['Weekly assignments (due Sunday)', 15, 30],
      ['Monday quizzes (lowest one dropped)', 12, 15],
      ['Exam 1 (weeks 1–5)', 1, 15],
      ['Exam 2 (weeks 6–10)', 1, 15],
      ['Exam 3 (weeks 11–15)', 1, 15],
      ['Final dashboard project and presentation', 1, 10],
    ],
    tools: [
      ['Microsoft Excel (Microsoft 365 desktop)', 'Weeks 1–15', 'Excel for the web covers weeks 1–10; full Power Query needs desktop Excel on Windows.'],
      ['The course workbook (this site)', 'Weeks 1–15', 'Runs in any browser, including Chromebooks. Every practice workbook downloads as a real .xlsx.'],
      ['Power Query', 'Preview in weeks 4 and 8; focus in weeks 11–15', 'Built into Excel and Power BI. This site includes a Power Query editor for Chromebook users.'],
      ['Power BI Desktop (free)', 'Weeks 13–15', 'Windows only. Chromebook students use the Power BI service in the browser with a school account, or a lab PC.'],
      ['A plain text editor (Notepad)', 'Week 4', 'To see that a CSV is just text and delimiters.'],
    ],
  };

  // ---------------------------------------------------------------------------
  // Seeded random numbers (mulberry32) so generated data never changes
  // ---------------------------------------------------------------------------
  function rng(seed) {
    let a = seed >>> 0;
    return () => {
      a |= 0; a = (a + 0x6d2b79f5) | 0;
      let t = Math.imul(a ^ (a >>> 15), 1 | a);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }
  const pick = (r, xs) => xs[Math.floor(r() * xs.length)];

  // ---------------------------------------------------------------------------
  // Shared datasets
  // ---------------------------------------------------------------------------
  const D = (C.data = {});

  // School store catalog (used weeks 6–14)
  D.catalog = [
    ['P101', 'Granola Bar', 'Snacks', 1.25],
    ['P102', 'Trail Mix', 'Snacks', 2.0],
    ['P103', 'Pretzels', 'Snacks', 1.5],
    ['P201', 'Water Bottle', 'Drinks', 1.0],
    ['P202', 'Sports Drink', 'Drinks', 2.25],
    ['P203', 'Iced Tea', 'Drinks', 1.75],
    ['P301', 'Notebook', 'Supplies', 3.5],
    ['P302', 'Pencil 4-Pack', 'Supplies', 2.0],
    ['P303', 'Graph Paper', 'Supplies', 4.0],
    ['P401', 'School T-Shirt', 'Apparel', 15.0],
    ['P402', 'Hoodie', 'Apparel', 32.0],
    ['P403', 'Lanyard', 'Apparel', 5.0],
  ];

  // 60 sales from January 5 to May 4, 2026 (weekdays only)
  D.storeSales = (() => {
    const r = rng(20260105);
    const rows = [];
    let serial = F.dateToSerial(2026, 1, 5);
    const cashiers = ['Avery', 'Jordan', 'Sam', 'Riley'];
    let id = 1001;
    while (rows.length < 60) {
      const dow = F.serialToDate(serial).dow;
      if (dow !== 0 && dow !== 6 && r() < 0.62) {
        const p = pick(r, D.catalog);
        const qty = p[2] === 'Apparel' ? 1 + Math.floor(r() * 2) : 1 + Math.floor(r() * 6);
        rows.push([id++, serial, p[0], p[1], p[2], qty, p[3], r() < 0.55 ? 'Card' : 'Cash', pick(r, cashiers)]);
      }
      serial++;
    }
    return rows;
  })();
  D.storeSalesHeader = ['OrderID', 'Date', 'ProductID', 'Item', 'Category', 'Qty', 'Price', 'Payment', 'Cashier'];

  // A sheet spec for the store sales data with optional extra columns.
  D.storeSalesSheet = (name = 'Sales', opts = {}) => {
    const header = [...D.storeSalesHeader, ...(opts.extraHeaders || [])];
    const rows = D.storeSales.slice(0, opts.limit || D.storeSales.length).map((r) => [...r]);
    const n = rows.length;
    return {
      name,
      rows: [header, ...rows],
      formats: { ['B2:B' + (n + 1)]: 'date', ['G2:G' + (n + 1)]: 'currency', ...(opts.formats || {}) },
      widths: { A: 70, B: 90, C: 80, D: 120, E: 85, F: 50, G: 70, H: 75, I: 75, ...(opts.widths || {}) },
      lock: opts.lock === undefined ? 'A1:I' + (n + 1) : opts.lock,
    };
  };

  // Week 1 class survey (15 students)
  D.survey = [
    ['Student', 'Grade', 'Favorite Lunch', 'Hours of Sleep', 'Has a Job', 'Minutes to School'],
    ['Maya', 12, 'Pizza', 7, 'Yes', 12],
    ['Luis', 11, 'Tacos', 6, 'No', 25],
    ['Aisha', 12, 'Salad', 8, 'No', 8],
    ['Ethan', 12, 'Pizza', 6.5, 'Yes', 15],
    ['Grace', 11, 'Burger', 7, 'No', 30],
    ['Noah', 12, 'Pizza', 5, 'Yes', 10],
    ['Sofia', 12, 'Tacos', 8, 'No', 18],
    ['Kai', 11, 'Pizza', 7.5, 'No', 22],
    ['Zoe', 12, 'Salad', 9, 'Yes', 5],
    ['Omar', 12, 'Burger', 6, 'Yes', 40],
    ['Lily', 11, 'Tacos', 7, 'No', 14],
    ['Ben', 12, 'Pizza', 6, 'No', 9],
    ['Chloe', 12, 'Salad', 8.5, 'Yes', 20],
    ['Diego', 11, 'Tacos', 7, 'Yes', 16],
    ['Hana', 12, 'Burger', 6.5, 'No', 11],
  ];

  // Students / Grades / Courses for relational thinking (weeks 7, 12, 13)
  D.students = [
    ['StudentID', 'First', 'Last', 'Grade', 'Advisor'],
    ['S1001', 'Maya', 'Patel', 12, 'Ms. Ortiz'],
    ['S1002', 'Luis', 'Garcia', 11, 'Mr. Chen'],
    ['S1003', 'Aisha', 'Bello', 12, 'Ms. Ortiz'],
    ['S1004', 'Ethan', 'Nguyen', 12, 'Mr. Chen'],
    ['S1005', 'Grace', 'Kim', 11, 'Ms. Hall'],
    ['S1006', 'Noah', 'Smith', 12, 'Ms. Hall'],
    ['S1007', 'Sofia', 'Rossi', 12, 'Mr. Chen'],
    ['S1008', 'Kai', 'Tanaka', 11, 'Ms. Ortiz'],
    ['S1009', 'Zoe', 'Martin', 12, 'Ms. Hall'],
    ['S1010', 'Omar', 'Haddad', 12, 'Mr. Chen'],
  ];
  D.courses = [
    ['CourseID', 'CourseName', 'Teacher', 'Credits'],
    ['C10', 'Data Management', 'Ms. Ortiz', 1],
    ['C20', 'Statistics', 'Mr. Chen', 1],
    ['C30', 'English 12', 'Ms. Hall', 1],
    ['C40', 'Web Design', 'Mr. Reyes', 0.5],
  ];
  D.grades = [
    ['EnrollmentID', 'StudentID', 'CourseID', 'Score'],
    ['E01', 'S1001', 'C10', 94], ['E02', 'S1001', 'C20', 88], ['E03', 'S1002', 'C10', 76],
    ['E04', 'S1002', 'C30', 85], ['E05', 'S1003', 'C20', 91], ['E06', 'S1004', 'C10', 67],
    ['E07', 'S1004', 'C40', 82], ['E08', 'S1005', 'C30', 95], ['E09', 'S1006', 'C10', 73],
    ['E10', 'S1007', 'C20', 89], ['E11', 'S1008', 'C10', 81], ['E12', 'S1009', 'C30', 98],
    ['E13', 'S1010', 'C40', 70], ['E14', 'S1012', 'C10', 84], ['E15', 'S1003', 'C10', 90],
    ['E16', 'S1011', 'C20', 77],
  ];

  // Monthly store exports for Power Query (weeks 11–12). Deliberately messy.
  function monthCsv(month, seed, n) {
    const r = rng(seed);
    const lines = ['Order ID,Order Date,Product ID,Qty,Unit Price,Store'];
    const stores = ['Main Office', 'main office ', 'Cafeteria', ' CAFETERIA', 'Gym Kiosk', 'gym kiosk'];
    for (let i = 0; i < n; i++) {
      const p = pick(r, D.catalog);
      const day = 1 + Math.floor(r() * 27);
      const qty = 1 + Math.floor(r() * 5);
      const price = p[3].toFixed(2);
      const date = month + '/' + day + '/2026';
      lines.push([month * 1000 + i + 1, date, p[0], qty, '$' + price, pick(r, stores)].join(','));
      if (i % 9 === 4) lines.push(lines[lines.length - 1]); // an exported duplicate
      if (i % 11 === 7) lines.push(',,,,,'); // a blank row from the export
    }
    return lines.join('\n');
  }
  D.files = {
    'sales_jan.csv': monthCsv(1, 101, 18),
    'sales_feb.csv': monthCsv(2, 202, 16),
    'sales_mar.csv': monthCsv(3, 303, 20),
    'products.csv': ['ProductID,ProductName,Category,Supplier'].concat(D.catalog.map((p) => [p[0], p[1], p[2], p[2] === 'Apparel' ? 'Spirit Wear Co' : p[2] === 'Supplies' ? 'OfficeHub' : 'SnackWorks'].join(','))).join('\n'),
    'attendance_wide.csv': [
      'Club,Sep,Oct,Nov,Dec',
      'Robotics,18,21,24,20',
      'Debate,12,15,14,11',
      'Art Club,22,19,25,23',
      'Coding Club,15,17,,19',
    ].join('\n'),
  };

  // Delimited files for week 4
  D.files['roster_comma.csv'] = 'StudentID,Name,Zip,Phone\n00417,"Patel, Maya",02134,555-0141\n00418,"Garcia, Luis",02139,555-0177\n00419,"Bello, Aisha",01002,555-0112\n00420,"Nguyen, Ethan",02215,555-0190';
  D.files['roster_tab.txt'] = 'StudentID\tName\tZip\tPhone\n00417\tMaya Patel\t02134\t555-0141\n00418\tLuis Garcia\t02139\t555-0177\n00419\tAisha Bello\t01002\t555-0112\n00420\tEthan Nguyen\t02215\t555-0190';
  D.files['roster_pipe.txt'] = 'StudentID|Name|Zip|Phone\n00417|Maya Patel|02134|555-0141\n00418|Luis Garcia|02139|555-0177\n00419|Aisha Bello|01002|555-0112\n00420|Ethan Nguyen|02215|555-0190';

  C.rng = rng;
  if (isNode) module.exports = C;
})(typeof globalThis !== 'undefined' ? globalThis : this);
