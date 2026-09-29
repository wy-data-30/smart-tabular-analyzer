// Fixed compatibility corpus. expected is the written calendar date, not a
// timezone-converted instant. Null cases intentionally reject loose Date.parse.
module.exports = [
  {
    "input": "Jan 1, 2025",
    "expected": "2025-01-01"
  },
  {
    "input": "Jan 1 2025",
    "expected": "2025-01-01"
  },
  {
    "input": "Jan 1, 2025 12:30:00",
    "expected": "2025-01-01"
  },
  {
    "input": "Jan 1, 2025, 12:30:00",
    "expected": "2025-01-01"
  },
  {
    "input": "Jan. 1, 2025 12:30 PM",
    "expected": "2025-01-01"
  },
  {
    "input": "Wed, 01 Jan 2025 12:30:00 GMT",
    "expected": "2025-01-01"
  },
  {
    "input": "Wed 01 Jan 2025 12:30:00 GMT",
    "expected": "2025-01-01"
  },
  {
    "input": "Wed Jan 1 12:30:00 2025",
    "expected": "2025-01-01"
  },
  {
    "input": "Wed Jan 01 00:00:00 2025",
    "expected": "2025-01-01"
  },
  {
    "input": "Wed Jan 01 2025 00:00:00 GMT+0800",
    "expected": "2025-01-01"
  },
  {
    "input": "1 Jan 2025",
    "expected": "2025-01-01"
  },
  {
    "input": "1-Jan-2025",
    "expected": "2025-01-01"
  },
  {
    "input": "1 Jan 2025 12:30:00",
    "expected": "2025-01-01"
  },
  {
    "input": "2025-Jan-1",
    "expected": "2025-01-01"
  },
  {
    "input": "2025 Jan 1",
    "expected": "2025-01-01"
  },
  {
    "input": "Jan 1, 2025 12:30:00 GMT",
    "expected": "2025-01-01"
  },
  {
    "input": "Jan 1, 2025 12:30:00 UTC",
    "expected": "2025-01-01"
  },
  {
    "input": "Jan 1, 2025 12:30:00 Z",
    "expected": "2025-01-01"
  },
  {
    "input": "Jan 1, 2025 12:30:00 +0800",
    "expected": "2025-01-01"
  },
  {
    "input": "Jan 1, 2025 12:30:00 +08:00",
    "expected": "2025-01-01"
  },
  {
    "input": "Jan 1, 2025 12:30:00 GMT+08:00",
    "expected": "2025-01-01"
  },
  {
    "input": "Jan 1, 2025 12:30:00 EST",
    "expected": "2025-01-01"
  },
  {
    "input": "Jan 1, 2025 12:30:00 PST",
    "expected": "2025-01-01"
  },
  {
    "input": "Feb 30, 2025",
    "expected": null,
    "reason": "invalid-calendar"
  },
  {
    "input": "Feb 29, 2025",
    "expected": null,
    "reason": "invalid-calendar"
  },
  {
    "input": "Apr 31, 2025",
    "expected": null,
    "reason": "invalid-calendar"
  },
  {
    "input": "Jan 32, 2025",
    "expected": null,
    "reason": "invalid-calendar"
  },
  {
    "input": "Jan 0, 2025",
    "expected": null,
    "reason": "invalid-calendar"
  },
  {
    "input": "29 Feb 1900",
    "expected": null,
    "reason": "invalid-calendar"
  },
  {
    "input": "Feb 29, 2024",
    "expected": "2024-02-29"
  },
  {
    "input": "29 Feb 2000",
    "expected": "2000-02-29"
  },
  {
    "input": "Dec 31, 2025 23:59:59",
    "expected": "2025-12-31"
  },
  {
    "input": "Dec 31, 2025 23:59:59 -1200",
    "expected": "2025-12-31"
  },
  {
    "input": "Jan 1, 2025 00:00:00 +1400",
    "expected": "2025-01-01"
  },
  {
    "input": "Wed Jan 01 2025 00:00:00 GMT+0800 (China Standard Time)",
    "expected": "2025-01-01"
  },
  {
    "input": "Jan 1, 2025 00:00:00.123Z",
    "expected": "2025-01-01"
  },
  {
    "input": "Jan 1, 25",
    "expected": null,
    "reason": "ambiguous-short-year"
  },
  {
    "input": "Januaryish 1, 2025",
    "expected": null,
    "reason": "invalid-month-name"
  },
  {
    "input": "Jan 1, 2025 24:00:00",
    "expected": null,
    "reason": "time-rollover"
  },
  {
    "input": "Jan 1, 2025 garbage",
    "expected": null,
    "reason": "unrecognized-suffix"
  }
];
