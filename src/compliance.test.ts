import { describe, it, expect } from 'vitest';

function detectMissingBreaks(shifts: { clockIn: string; clockOut: string; hours: number }[], longShiftThreshold = 5) {
  return shifts.filter(s => s.hours >= longShiftThreshold);
}

describe('Compliance Rule Engine', () => {
  it('flags shifts exceeding long shift threshold without recorded break', () => {
    const shifts = [
      { clockIn: '2026-07-28T08:00:00Z', clockOut: '2026-07-28T16:00:00Z', hours: 8.0 },
      { clockIn: '2026-07-28T08:00:00Z', clockOut: '2026-07-28T12:00:00Z', hours: 4.0 },
    ];

    const flagged = detectMissingBreaks(shifts, 5);
    expect(flagged.length).toBe(1);
    expect(flagged[0].hours).toBe(8.0);
  });
});
