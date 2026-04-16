// ==================== Google Sheets API Wrapper ====================

const Sheets = (() => {

  async function _request(fn) {
    await Auth.ensureToken();
    try {
      const result = await fn();
      return result;
    } catch (e) {
      const msg = e?.result?.error?.message || e?.message || 'Sheets API error';
      console.error('Sheets error:', msg, e);
      throw new Error(msg);
    }
  }

  async function getRange(sheet, range) {
    return _request(async () => {
      const r = await gapi.client.sheets.spreadsheets.values.get({
        spreadsheetId: CONFIG.SPREADSHEET_ID,
        range: range ? `${sheet}!${range}` : sheet
      });
      return r.result.values || [];
    });
  }

  async function appendRows(sheet, values) {
    return _request(async () => {
      const r = await gapi.client.sheets.spreadsheets.values.append({
        spreadsheetId: CONFIG.SPREADSHEET_ID,
        range: `${sheet}!A1`,
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        resource: { values }
      });
      return r.result;
    });
  }

  async function updateRow(sheet, rowIndex, values) {
    // rowIndex is 1-based (1 = header, 2 = first data row)
    return _request(async () => {
      const r = await gapi.client.sheets.spreadsheets.values.update({
        spreadsheetId: CONFIG.SPREADSHEET_ID,
        range: `${sheet}!A${rowIndex}`,
        valueInputOption: 'USER_ENTERED',
        resource: { values: [values] }
      });
      return r.result;
    });
  }

  async function deleteRow(sheet, rowIndex) {
    return _request(async () => {
      // Get the sheet ID first
      const meta = await gapi.client.sheets.spreadsheets.get({
        spreadsheetId: CONFIG.SPREADSHEET_ID
      });
      const sheetMeta = meta.result.sheets.find(s => s.properties.title === sheet);
      if (!sheetMeta) throw new Error(`Sheet ${sheet} not found`);

      const r = await gapi.client.sheets.spreadsheets.batchUpdate({
        spreadsheetId: CONFIG.SPREADSHEET_ID,
        resource: {
          requests: [{
            deleteDimension: {
              range: {
                sheetId: sheetMeta.properties.sheetId,
                dimension: 'ROWS',
                startIndex: rowIndex - 1,
                endIndex: rowIndex
              }
            }
          }]
        }
      });
      return r.result;
    });
  }

  async function clearRange(sheet, range) {
    return _request(async () => {
      const r = await gapi.client.sheets.spreadsheets.values.clear({
        spreadsheetId: CONFIG.SPREADSHEET_ID,
        range: range ? `${sheet}!${range}` : sheet
      });
      return r.result;
    });
  }

  async function batchUpdate(sheet, updates) {
    // updates: [{rowIndex, values}]
    return _request(async () => {
      const data = updates.map(u => ({
        range: `${sheet}!A${u.rowIndex}`,
        values: [u.values]
      }));
      const r = await gapi.client.sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: CONFIG.SPREADSHEET_ID,
        resource: { valueInputOption: 'USER_ENTERED', data }
      });
      return r.result;
    });
  }

  async function ensureSheetExists(sheetName, headers) {
    try {
      const meta = await gapi.client.sheets.spreadsheets.get({
        spreadsheetId: CONFIG.SPREADSHEET_ID
      });
      const exists = meta.result.sheets.some(s => s.properties.title === sheetName);
      if (!exists) {
        await gapi.client.sheets.spreadsheets.batchUpdate({
          spreadsheetId: CONFIG.SPREADSHEET_ID,
          resource: {
            requests: [{ addSheet: { properties: { title: sheetName } } }]
          }
        });
        await appendRows(sheetName, [headers]);
      }
    } catch (e) {
      console.warn('ensureSheetExists error:', e);
    }
  }

  async function initSheets() {
    const sheetsConfig = [
      {
        name: CONFIG.SHEET_NAMES.MEMBERS,
        headers: ['member_id','name','id_card','phone','address','join_date','status']
      },
      {
        name: CONFIG.SHEET_NAMES.DEPOSITS,
        headers: ['transaction_id','member_id','member_name','type','amount','month','year','date','recorded_by','notes']
      },
      {
        name: CONFIG.SHEET_NAMES.SHARES,
        headers: ['share_id','member_id','member_name','shares_count','transaction_type','date','recorded_by']
      },
      {
        name: CONFIG.SHEET_NAMES.LOANS,
        headers: ['loan_id','member_id','member_name','loan_amount','loan_date','duration_months',
          'guarantor1_id','guarantor1_name','guarantor2_id','guarantor2_name','status',
          'drive_folder_id','contract_file','id_card_file','house_reg_file','photo_file']
      },
      {
        name: CONFIG.SHEET_NAMES.LOAN_PAYMENTS,
        headers: ['payment_id','loan_id','member_id','member_name','payment_date',
          'amount_paid','principal','interest','penalty','remaining_principal','month_number','is_late']
      },
      {
        name: CONFIG.SHEET_NAMES.COMMITTEE,
        headers: ['record_id','member_id','member_name','work_date','month','year','wage_per_session','recorded_by']
      },
      {
        name: CONFIG.SHEET_NAMES.INTEREST_CALC,
        headers: ['calc_id','member_id','member_name','year',
          'jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec',
          'total_deposit','total_interest','months_deposited']
      }
    ];

    for (const s of sheetsConfig) {
      await ensureSheetExists(s.name, s.headers);
    }
  }

  // ==================== MEMBERS ====================

  async function getMembers(useCache = true) {
    const cacheKey = `cb_${CONFIG.SHEET_NAMES.MEMBERS}`;
    if (useCache) {
      const cached = Cache.get(cacheKey);
      if (cached) return cached;
    }
    const rows = await getRange(CONFIG.SHEET_NAMES.MEMBERS, 'A2:G');
    const members = rows.map(r => ({
      member_id: r[0] || '',
      name: r[1] || '',
      id_card: r[2] || '',
      phone: r[3] || '',
      address: r[4] || '',
      join_date: r[5] || '',
      status: r[6] || 'active'
    }));
    Cache.set(cacheKey, members);
    return members;
  }

  async function addMember(member) {
    const members = await getMembers(false);
    const newId = generateMemberId(members);
    const row = [newId, member.name, member.id_card, member.phone,
      member.address, member.join_date, member.status || 'active'];
    await appendRows(CONFIG.SHEET_NAMES.MEMBERS, [row]);
    Cache.invalidate(CONFIG.SHEET_NAMES.MEMBERS);
    return { ...member, member_id: newId };
  }

  async function updateMember(member) {
    const rows = await getRange(CONFIG.SHEET_NAMES.MEMBERS, 'A2:G');
    const rowIndex = rows.findIndex(r => r[0] === member.member_id);
    if (rowIndex === -1) throw new Error('ไม่พบสมาชิก');
    const values = [member.member_id, member.name, member.id_card, member.phone,
      member.address, member.join_date, member.status];
    await updateRow(CONFIG.SHEET_NAMES.MEMBERS, rowIndex + 2, values);
    Cache.invalidate(CONFIG.SHEET_NAMES.MEMBERS);
  }

  async function deleteMember(memberId) {
    const rows = await getRange(CONFIG.SHEET_NAMES.MEMBERS, 'A2:G');
    const rowIndex = rows.findIndex(r => r[0] === memberId);
    if (rowIndex === -1) throw new Error('ไม่พบสมาชิก');
    await deleteRow(CONFIG.SHEET_NAMES.MEMBERS, rowIndex + 2);
    Cache.invalidate(CONFIG.SHEET_NAMES.MEMBERS);
  }

  // ==================== DEPOSITS ====================

  async function getDeposits(useCache = true) {
    const cacheKey = `cb_${CONFIG.SHEET_NAMES.DEPOSITS}`;
    if (useCache) {
      const cached = Cache.get(cacheKey);
      if (cached) return cached;
    }
    const rows = await getRange(CONFIG.SHEET_NAMES.DEPOSITS, 'A2:J');
    const deposits = rows.map(r => ({
      transaction_id: r[0] || '',
      member_id: r[1] || '',
      member_name: r[2] || '',
      type: r[3] || '',
      amount: parseFloat(r[4]) || 0,
      month: parseInt(r[5]) || 0,
      year: parseInt(r[6]) || 0,
      date: r[7] || '',
      recorded_by: r[8] || '',
      notes: r[9] || ''
    }));
    Cache.set(cacheKey, deposits);
    return deposits;
  }

  async function addDeposit(deposit) {
    const deposits = await getDeposits(false);
    const newId = generateId('TXN', deposits, 'transaction_id');
    const user = Auth.getUser();
    const row = [newId, deposit.member_id, deposit.member_name, deposit.type,
      deposit.amount, deposit.month, deposit.year, deposit.date,
      user ? user.name : 'System', deposit.notes || ''];
    await appendRows(CONFIG.SHEET_NAMES.DEPOSITS, [row]);
    Cache.invalidate(CONFIG.SHEET_NAMES.DEPOSITS);
    return { ...deposit, transaction_id: newId };
  }

  async function deleteDeposit(transactionId) {
    const rows = await getRange(CONFIG.SHEET_NAMES.DEPOSITS, 'A2:J');
    const rowIndex = rows.findIndex(r => r[0] === transactionId);
    if (rowIndex === -1) throw new Error('ไม่พบรายการ');
    await deleteRow(CONFIG.SHEET_NAMES.DEPOSITS, rowIndex + 2);
    Cache.invalidate(CONFIG.SHEET_NAMES.DEPOSITS);
  }

  // ==================== SHARES ====================

  async function getShares(useCache = true) {
    const cacheKey = `cb_${CONFIG.SHEET_NAMES.SHARES}`;
    if (useCache) {
      const cached = Cache.get(cacheKey);
      if (cached) return cached;
    }
    const rows = await getRange(CONFIG.SHEET_NAMES.SHARES, 'A2:G');
    const shares = rows.map(r => ({
      share_id: r[0] || '',
      member_id: r[1] || '',
      member_name: r[2] || '',
      shares_count: parseInt(r[3]) || 0,
      transaction_type: r[4] || '',
      date: r[5] || '',
      recorded_by: r[6] || ''
    }));
    Cache.set(cacheKey, shares);
    return shares;
  }

  async function addShareTransaction(tx) {
    const shares = await getShares(false);
    const newId = generateId('SH', shares, 'share_id');
    const user = Auth.getUser();
    const row = [newId, tx.member_id, tx.member_name, tx.shares_count,
      tx.transaction_type, tx.date, user ? user.name : 'System'];
    await appendRows(CONFIG.SHEET_NAMES.SHARES, [row]);
    Cache.invalidate(CONFIG.SHEET_NAMES.SHARES);
    return { ...tx, share_id: newId };
  }

  function getMemberShareBalance(allShares, memberId) {
    return allShares
      .filter(s => s.member_id === memberId)
      .reduce((sum, s) => {
        return sum + (s.transaction_type === 'buy' ? s.shares_count : -s.shares_count);
      }, 0);
  }

  // ==================== LOANS ====================

  async function getLoans(useCache = true) {
    const cacheKey = `cb_${CONFIG.SHEET_NAMES.LOANS}`;
    if (useCache) {
      const cached = Cache.get(cacheKey);
      if (cached) return cached;
    }
    const rows = await getRange(CONFIG.SHEET_NAMES.LOANS, 'A2:P');
    const loans = rows.map(r => ({
      loan_id: r[0] || '',
      member_id: r[1] || '',
      member_name: r[2] || '',
      loan_amount: parseFloat(r[3]) || 0,
      loan_date: r[4] || '',
      duration_months: parseInt(r[5]) || 0,
      guarantor1_id: r[6] || '',
      guarantor1_name: r[7] || '',
      guarantor2_id: r[8] || '',
      guarantor2_name: r[9] || '',
      status: r[10] || 'active',
      drive_folder_id: r[11] || '',
      contract_file: r[12] || '',
      id_card_file: r[13] || '',
      house_reg_file: r[14] || '',
      photo_file: r[15] || ''
    }));
    Cache.set(cacheKey, loans);
    return loans;
  }

  async function addLoan(loan) {
    const loans = await getLoans(false);
    const newId = generateLoanId(loans);
    const row = [newId, loan.member_id, loan.member_name, loan.loan_amount, loan.loan_date,
      loan.duration_months, loan.guarantor1_id || '', loan.guarantor1_name || '',
      loan.guarantor2_id || '', loan.guarantor2_name || '', 'active',
      loan.drive_folder_id || '', loan.contract_file || '', loan.id_card_file || '',
      loan.house_reg_file || '', loan.photo_file || ''];
    await appendRows(CONFIG.SHEET_NAMES.LOANS, [row]);
    Cache.invalidate(CONFIG.SHEET_NAMES.LOANS);
    return { ...loan, loan_id: newId, status: 'active' };
  }

  async function updateLoan(loan) {
    const rows = await getRange(CONFIG.SHEET_NAMES.LOANS, 'A2:P');
    const rowIndex = rows.findIndex(r => r[0] === loan.loan_id);
    if (rowIndex === -1) throw new Error('ไม่พบข้อมูลเงินกู้');
    const values = [loan.loan_id, loan.member_id, loan.member_name, loan.loan_amount,
      loan.loan_date, loan.duration_months, loan.guarantor1_id || '', loan.guarantor1_name || '',
      loan.guarantor2_id || '', loan.guarantor2_name || '', loan.status,
      loan.drive_folder_id || '', loan.contract_file || '', loan.id_card_file || '',
      loan.house_reg_file || '', loan.photo_file || ''];
    await updateRow(CONFIG.SHEET_NAMES.LOANS, rowIndex + 2, values);
    Cache.invalidate(CONFIG.SHEET_NAMES.LOANS);
  }

  // ==================== LOAN PAYMENTS ====================

  async function getLoanPayments(useCache = true) {
    const cacheKey = `cb_${CONFIG.SHEET_NAMES.LOAN_PAYMENTS}`;
    if (useCache) {
      const cached = Cache.get(cacheKey);
      if (cached) return cached;
    }
    const rows = await getRange(CONFIG.SHEET_NAMES.LOAN_PAYMENTS, 'A2:L');
    const payments = rows.map(r => ({
      payment_id: r[0] || '',
      loan_id: r[1] || '',
      member_id: r[2] || '',
      member_name: r[3] || '',
      payment_date: r[4] || '',
      amount_paid: parseFloat(r[5]) || 0,
      principal: parseFloat(r[6]) || 0,
      interest: parseFloat(r[7]) || 0,
      penalty: parseFloat(r[8]) || 0,
      remaining_principal: parseFloat(r[9]) || 0,
      month_number: parseInt(r[10]) || 0,
      is_late: r[11] === 'TRUE'
    }));
    Cache.set(cacheKey, payments);
    return payments;
  }

  async function addLoanPayment(payment) {
    const payments = await getLoanPayments(false);
    const newId = generateId('PMT', payments, 'payment_id');
    const row = [newId, payment.loan_id, payment.member_id, payment.member_name,
      payment.payment_date, payment.amount_paid, payment.principal, payment.interest,
      payment.penalty || 0, payment.remaining_principal, payment.month_number,
      payment.is_late ? 'TRUE' : 'FALSE'];
    await appendRows(CONFIG.SHEET_NAMES.LOAN_PAYMENTS, [row]);
    Cache.invalidate(CONFIG.SHEET_NAMES.LOAN_PAYMENTS);
    return { ...payment, payment_id: newId };
  }

  function getLoanRemainingBalance(allPayments, loanId, loanAmount) {
    const loanPayments = allPayments
      .filter(p => p.loan_id === loanId)
      .sort((a, b) => b.month_number - a.month_number);
    if (loanPayments.length === 0) return loanAmount;
    return loanPayments[0].remaining_principal;
  }

  function getNextPaymentInfo(allPayments, loan) {
    const loanPayments = allPayments.filter(p => p.loan_id === loan.loan_id);
    const monthNumber = loanPayments.length + 1;
    const balance = getLoanRemainingBalance(allPayments, loan.loan_id, loan.loan_amount);
    const rate = CONFIG.INTEREST.LOAN_MONTHLY_RATE / 100;

    // Check if previous payment was late
    const lastPayment = loanPayments.sort((a, b) => b.month_number - a.month_number)[0];
    const wasLate = lastPayment && lastPayment.is_late;

    const interestMultiplier = wasLate ? 2 : 1;
    const interest = balance * rate * interestMultiplier;
    const penalty = wasLate ? CONFIG.INTEREST.LATE_PENALTY : 0;

    const monthlyPayment = calculateLoanPayment(
      loan.loan_amount, CONFIG.INTEREST.LOAN_MONTHLY_RATE, loan.duration_months
    );
    const principal = Math.min(monthlyPayment - (balance * rate), balance);

    return {
      month_number: monthNumber,
      balance,
      interest,
      penalty,
      principal: Math.max(0, principal),
      suggested_amount: interest + principal + penalty,
      was_late: wasLate
    };
  }

  // ==================== COMMITTEE ====================

  async function getCommittee(useCache = true) {
    const cacheKey = `cb_${CONFIG.SHEET_NAMES.COMMITTEE}`;
    if (useCache) {
      const cached = Cache.get(cacheKey);
      if (cached) return cached;
    }
    const rows = await getRange(CONFIG.SHEET_NAMES.COMMITTEE, 'A2:H');
    const records = rows.map(r => ({
      record_id: r[0] || '',
      member_id: r[1] || '',
      member_name: r[2] || '',
      work_date: r[3] || '',
      month: parseInt(r[4]) || 0,
      year: parseInt(r[5]) || 0,
      wage_per_session: parseFloat(r[6]) || 0,
      recorded_by: r[7] || ''
    }));
    Cache.set(cacheKey, records);
    return records;
  }

  async function addCommitteeRecord(record) {
    const records = await getCommittee(false);
    const newId = generateId('COM', records, 'record_id');
    const user = Auth.getUser();
    const row = [newId, record.member_id, record.member_name, record.work_date,
      record.month, record.year, record.wage_per_session, user ? user.name : 'System'];
    await appendRows(CONFIG.SHEET_NAMES.COMMITTEE, [row]);
    Cache.invalidate(CONFIG.SHEET_NAMES.COMMITTEE);
    return { ...record, record_id: newId };
  }

  async function deleteCommitteeRecord(recordId) {
    const rows = await getRange(CONFIG.SHEET_NAMES.COMMITTEE, 'A2:H');
    const rowIndex = rows.findIndex(r => r[0] === recordId);
    if (rowIndex === -1) throw new Error('ไม่พบรายการ');
    await deleteRow(CONFIG.SHEET_NAMES.COMMITTEE, rowIndex + 2);
    Cache.invalidate(CONFIG.SHEET_NAMES.COMMITTEE);
  }

  // ==================== INTEREST CALC ====================

  async function getInterestCalc(useCache = true) {
    const cacheKey = `cb_${CONFIG.SHEET_NAMES.INTEREST_CALC}`;
    if (useCache) {
      const cached = Cache.get(cacheKey);
      if (cached) return cached;
    }
    const rows = await getRange(CONFIG.SHEET_NAMES.INTEREST_CALC, 'A2:S');
    const calcs = rows.map(r => ({
      calc_id: r[0] || '',
      member_id: r[1] || '',
      member_name: r[2] || '',
      year: parseInt(r[3]) || 0,
      jan: parseFloat(r[4]) || 0,
      feb: parseFloat(r[5]) || 0,
      mar: parseFloat(r[6]) || 0,
      apr: parseFloat(r[7]) || 0,
      may: parseFloat(r[8]) || 0,
      jun: parseFloat(r[9]) || 0,
      jul: parseFloat(r[10]) || 0,
      aug: parseFloat(r[11]) || 0,
      sep: parseFloat(r[12]) || 0,
      oct: parseFloat(r[13]) || 0,
      nov: parseFloat(r[14]) || 0,
      dec: parseFloat(r[15]) || 0,
      total_deposit: parseFloat(r[16]) || 0,
      total_interest: parseFloat(r[17]) || 0,
      months_deposited: parseInt(r[18]) || 0
    }));
    Cache.set(cacheKey, calcs);
    return calcs;
  }

  async function saveInterestCalc(calc) {
    const calcs = await getInterestCalc(false);
    const existing = calcs.find(c => c.member_id === calc.member_id && c.year === calc.year);
    const monthKeys = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
    const row = [
      calc.calc_id || generateId('IC', calcs, 'calc_id'),
      calc.member_id, calc.member_name, calc.year,
      ...monthKeys.map(m => calc[m] || 0),
      calc.total_deposit, calc.total_interest, calc.months_deposited
    ];

    if (existing) {
      const rows = await getRange(CONFIG.SHEET_NAMES.INTEREST_CALC, 'A2:S');
      const rowIndex = rows.findIndex(r => r[1] === calc.member_id && parseInt(r[3]) === calc.year);
      if (rowIndex !== -1) {
        await updateRow(CONFIG.SHEET_NAMES.INTEREST_CALC, rowIndex + 2, row);
      }
    } else {
      await appendRows(CONFIG.SHEET_NAMES.INTEREST_CALC, [row]);
    }
    Cache.invalidate(CONFIG.SHEET_NAMES.INTEREST_CALC);
  }

  // ==================== Business Logic: Interest Calculation ====================

  function calculateMemberInterest(memberId, year, deposits) {
    const monthKeys = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
    const memberDeposits = deposits.filter(d => d.member_id === memberId && d.year === year);

    // Running balance per month (net deposits)
    const monthlyBalance = {};
    let runningBalance = 0;

    // Get opening balance (all deposits before this year)
    const previousDeposits = deposits.filter(d => d.member_id === memberId && d.year < year);
    previousDeposits.forEach(d => {
      runningBalance += d.type === 'deposit' ? d.amount : -d.amount;
    });

    for (let m = 1; m <= 12; m++) {
      const monthDeps = memberDeposits.filter(d => d.month === m);
      monthDeps.forEach(d => {
        runningBalance += d.type === 'deposit' ? d.amount : -d.amount;
      });
      monthlyBalance[m] = Math.max(0, runningBalance);
    }

    // Calculate interest: deposit in month m earns rate for that month
    let totalInterest = 0;
    let monthsDeposited = 0;
    const monthInterests = {};

    for (let m = 1; m <= 12; m++) {
      const balance = monthlyBalance[m];
      if (balance > 0) {
        const rate = getDepositInterestRate(m) / 100;
        const interest = balance * rate;
        monthInterests[monthKeys[m - 1]] = interest;
        totalInterest += interest;
        monthsDeposited++;
      } else {
        monthInterests[monthKeys[m - 1]] = 0;
      }
    }

    const totalDeposit = monthlyBalance[12] || 0;

    return {
      member_id: memberId,
      year,
      ...monthInterests,
      total_deposit: totalDeposit,
      total_interest: totalInterest,
      months_deposited: monthsDeposited
    };
  }

  return {
    initSheets,
    getMembers, addMember, updateMember, deleteMember,
    getDeposits, addDeposit, deleteDeposit,
    getShares, addShareTransaction, getMemberShareBalance,
    getLoans, addLoan, updateLoan,
    getLoanPayments, addLoanPayment, getLoanRemainingBalance, getNextPaymentInfo,
    getCommittee, addCommitteeRecord, deleteCommitteeRecord,
    getInterestCalc, saveInterestCalc, calculateMemberInterest
  };
})();
