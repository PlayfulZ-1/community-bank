const CONFIG = {
  CLIENT_ID: '372266975424-dlv6d7j7br0gens980l589jcotev8nj8.apps.googleusercontent.com',
  SPREADSHEET_ID: '1ZHJgkd1UsMdBIBHIIuJRr2gFlSlh3Xj1n15piNaynA8',
  DRIVE_FOLDER_ID: '1UfajMdRMpt1PjGuVUuhvCc2o-3yzGbWr',
  SCOPES: 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file',
  DISCOVERY_DOCS: [
    'https://sheets.googleapis.com/$discovery/rest?version=v4',
    'https://www.googleapis.com/discovery/v1/apis/drive/v3/rest'
  ],
  SHEET_NAMES: {
    MEMBERS: 'MEMBERS',
    DEPOSITS: 'DEPOSITS',
    SHARES: 'SHARES',
    LOANS: 'LOANS',
    LOAN_PAYMENTS: 'LOAN_PAYMENTS',
    COMMITTEE: 'COMMITTEE',
    INTEREST_CALC: 'INTEREST_CALC'
  },
  INTEREST: {
    DEPOSIT_BASE_RATE: 3.0,
    DEPOSIT_DECREMENT: 0.25,
    LOAN_MONTHLY_RATE: 1.0,
    LATE_PENALTY: 200
  },
  SHARE_RULES: {
    MIN_SHARES: 10,
    MAX_SHARES: 200,
    PRICE_PER_SHARE: 10
  },
  PROFIT_DISTRIBUTION: {
    CAPITAL_FUND_RATIO: 0.50,
    COMMITTEE_RATIO: 0.30,
    COMMUNITY_FUND_RATIO: 0.50,
    DIVIDEND_RATIO: 0.20
  },
  CACHE_TTL_MS: 5 * 60 * 1000
};
