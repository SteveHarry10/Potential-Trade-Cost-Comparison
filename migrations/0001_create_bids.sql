CREATE TABLE bids (
  id TEXT PRIMARY KEY NOT NULL,
  plan TEXT NOT NULL,
  code TEXT NOT NULL,
  trade TEXT NOT NULL,
  description TEXT NOT NULL,
  current REAL NOT NULL,
  bidder TEXT NOT NULL,
  amount REAL NOT NULL,
  bid_date TEXT NOT NULL,
  scope TEXT NOT NULL DEFAULT 'Base bid',
  notes TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX idx_bids_created_at ON bids(created_at DESC);
CREATE INDEX idx_bids_plan_code ON bids(plan, code);
