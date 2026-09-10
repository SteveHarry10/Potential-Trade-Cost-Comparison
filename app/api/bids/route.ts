import { env } from 'cloudflare:workers';
import { NextResponse } from 'next/server';

type BidInput = {
  plan?: unknown;
  code?: unknown;
  trade?: unknown;
  description?: unknown;
  current?: unknown;
  bidder?: unknown;
  amount?: unknown;
  date?: unknown;
  scope?: unknown;
  notes?: unknown;
};

type BidRow = {
  bidId: string;
  plan: string;
  code: string;
  trade: string;
  description: string;
  current: number;
  bidder: string;
  amount: number;
  date: string;
  scope: string;
  notes: string;
};

function database() {
  const db = (env as unknown as { DB?: D1Database }).DB;
  if (!db) throw new Error('The shared bids database is not connected. Add the DB binding in Cloudflare.');
  return db;
}

function text(value: unknown) {
  return String(value ?? '').trim();
}

function number(value: unknown) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : null;
}

export async function GET() {
  try {
    const result = await database().prepare(`
      SELECT
        id AS bidId,
        plan,
        code,
        trade,
        description,
        current,
        bidder,
        amount,
        bid_date AS date,
        scope,
        notes
      FROM bids
      ORDER BY created_at DESC
    `).all<BidRow>();
    return NextResponse.json({ bids: result.results ?? [] });
  } catch (error) {
    console.error('Unable to load shared bids', error);
    return NextResponse.json({ message: error instanceof Error ? error.message : 'Could not load shared bids.' }, { status: 503 });
  }
}

export async function POST(request: Request) {
  try {
    const input = await request.json() as BidInput;
    const current = number(input.current);
    const amount = number(input.amount);
    const bid: BidRow = {
      bidId: crypto.randomUUID(),
      plan: text(input.plan),
      code: text(input.code),
      trade: text(input.trade),
      description: text(input.description),
      current: current ?? 0,
      bidder: text(input.bidder),
      amount: amount ?? 0,
      date: text(input.date),
      scope: text(input.scope) || 'Base bid',
      notes: text(input.notes),
    };

    if (!bid.plan || !bid.code || !bid.bidder || !bid.date || current === null || amount === null) {
      return NextResponse.json({ message: 'The bid is missing a plan, cost code, bidder, date, or valid amount.' }, { status: 400 });
    }

    await database().prepare(`
      INSERT INTO bids (
        id, plan, code, trade, description, current, bidder, amount, bid_date, scope, notes
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).bind(
      bid.bidId,
      bid.plan,
      bid.code,
      bid.trade,
      bid.description,
      bid.current,
      bid.bidder,
      bid.amount,
      bid.date,
      bid.scope,
      bid.notes,
    ).run();

    return NextResponse.json({ bid }, { status: 201 });
  } catch (error) {
    console.error('Unable to save shared bid', error);
    return NextResponse.json({ message: error instanceof Error ? error.message : 'Could not save the shared bid.' }, { status: 503 });
  }
}

export async function DELETE(request: Request) {
  try {
    const bidId = new URL(request.url).searchParams.get('id')?.trim();
    if (!bidId) return NextResponse.json({ message: 'A bid ID is required.' }, { status: 400 });
    await database().prepare('DELETE FROM bids WHERE id = ?').bind(bidId).run();
    return NextResponse.json({ deleted: true });
  } catch (error) {
    console.error('Unable to delete shared bid', error);
    return NextResponse.json({ message: error instanceof Error ? error.message : 'Could not delete the shared bid.' }, { status: 503 });
  }
}
