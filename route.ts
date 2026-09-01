import { NextResponse } from 'next/server';

type IncomingRow = Record<string, unknown>;

function value(row: IncomingRow, ...keys: string[]) {
  const matched = keys.find((key) => row[key] !== undefined && row[key] !== null);
  return matched ? row[matched] : undefined;
}

function numberFrom(value: unknown) {
  const parsed = Number(String(value ?? '').replace(/[$,\s]/g, ''));
  return Number.isFinite(parsed) ? parsed : null;
}

function normalize(rows: IncomingRow[]) {
  const alreadyNormalized = rows.some((row) => value(row, 'plan', 'Plan', 'Floorplan') !== undefined);
  if (alreadyNormalized) {
    return rows.map((row, index) => ({
      id: Number(value(row, 'id', 'ID')) || index + 1,
      plan: String(value(row, 'plan', 'Plan', 'Floorplan') ?? '').trim(),
      code: String(value(row, 'code', 'Code', 'Cost Code') ?? '').trim(),
      trade: String(value(row, 'trade', 'Trade', 'Category') ?? '').trim(),
      description: String(value(row, 'description', 'Description', 'Item') ?? '').trim(),
      current: numberFrom(value(row, 'current', 'Current', 'Cost', 'Amount')),
    })).filter((row) => row.plan && row.code && row.current !== null);
  }

  const metadata = new Set(['', 'Code', 'Cost Code', 'Item', 'Description', '@odata.etag', 'ItemInternalId']);
  const costs: Array<{ id: number; plan: string; code: string; trade: string; description: string; current: number }> = [];
  rows.forEach((row) => {
    const code = String(value(row, 'Code', 'Cost Code', 'code') ?? '').trim();
    const description = String(value(row, 'Item', 'Description', 'description') ?? '').trim();
    if (!/^\d{3}-\d{2}$/.test(code) || !description) return;
    const trade = description.includes(' - ') ? description.split(' - ')[0].trim() : description;
    Object.entries(row).forEach(([planName, raw]) => {
      if (metadata.has(planName)) return;
      const current = numberFrom(raw);
      if (current === null || String(raw).trim().startsWith('#')) return;
      costs.push({ id: costs.length + 1, plan: planName.trim(), code, trade, description, current });
    });
  });
  return costs;
}

export async function GET() {
  const source = process.env.SHAREPOINT_EXCEL_ENDPOINT;
  if (!source) {
    return NextResponse.json({
      connected: false,
      message: 'Add SHAREPOINT_EXCEL_ENDPOINT to connect the Floorplans workbook. See README.md for the one-time setup.',
    });
  }

  try {
    const response = await fetch(source, {
      headers: process.env.SHAREPOINT_EXCEL_TOKEN ? { Authorization: `Bearer ${process.env.SHAREPOINT_EXCEL_TOKEN}` } : undefined,
      cache: 'no-store',
    });
    if (!response.ok) throw new Error(`Source returned ${response.status}`);
    const payload = await response.json() as IncomingRow[] | { value?: IncomingRow[]; rows?: IncomingRow[] };
    const rows = Array.isArray(payload) ? payload : payload.value ?? payload.rows ?? [];
    const costs = normalize(rows);
    return NextResponse.json({ connected: true, costs, updatedAt: new Date().toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' }) });
  } catch (error) {
    return NextResponse.json({ connected: false, message: error instanceof Error ? error.message : 'The SharePoint source could not be read.' }, { status: 502 });
  }
}
