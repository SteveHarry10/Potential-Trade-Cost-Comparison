'use client';

import { FormEvent, useEffect, useMemo, useState } from 'react';
import { ArrowDownRight, ArrowUpRight, Building2, Check, Cloud, FileSpreadsheet, Plus, RefreshCw, Search, Trash2 } from 'lucide-react';
import floorplansData from './data/floorplans-data.json';

type CostLine = { id: number; plan: string; code: string; trade: string; description: string; current: number };
type Bid = CostLine & { bidId: string; bidder: string; amount: number; date: string; scope: string; notes: string };
type FloorplansData = {
  source: { importedAt: string };
  items: Array<{ code: string; item: string; category: string }>;
  plans: Array<{ name: string; values: Array<number | null> }>;
};

const sourceData = floorplansData as FloorplansData;
const sourceCosts: CostLine[] = sourceData.plans.flatMap((sourcePlan, planIndex) =>
  sourceData.items.flatMap((item, itemIndex) => {
    const current = sourcePlan.values[itemIndex];
    return typeof current === 'number'
      ? [{ id: planIndex * 1000 + itemIndex + 1, plan: sourcePlan.name, code: item.code, trade: item.category, description: item.item, current }]
      : [];
  }),
);
const refreshWorkflowUrl = 'https://github.com/SteveHarry10/Potential-Trade-Cost-Comparison/actions/workflows/refresh-dashboard.yml';

const money = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 });
const dateLabel = new Intl.DateTimeFormat('en-US', { month: 'short', day: 'numeric', year: 'numeric' });

export default function Dashboard() {
  const [plan, setPlan] = useState('All plans');
  const [trade, setTrade] = useState('All trades');
  const [costRows] = useState<CostLine[]>(sourceCosts);
  const [bidPlan, setBidPlan] = useState(sourceCosts[0]?.plan ?? '');
  const [selectedId, setSelectedId] = useState(sourceCosts[0]?.id ?? 0);
  const [bidder, setBidder] = useState('');
  const [amount, setAmount] = useState('');
  const [bidDate, setBidDate] = useState(new Date().toISOString().slice(0, 10));
  const [scope, setScope] = useState('Base bid');
  const [notes, setNotes] = useState('');
  const [bids, setBids] = useState<Bid[]>([]);
  const [query, setQuery] = useState('');
  const [notice, setNotice] = useState('');
  const [saving, setSaving] = useState(false);
  const costStatus = 'GitHub synced';

  useEffect(() => {
    void loadBids();
  }, []);

  const plans = [...new Set(costRows.map((row) => row.plan))];
  const trades = [...new Set(costRows.map((row) => row.trade))];
  const bidCostOptions = costRows.filter((row) => row.plan === bidPlan);
  const selected = costRows.find((row) => row.id === selectedId && row.plan === bidPlan) ?? bidCostOptions[0] ?? costRows[0];
  const visibleBids = bids.map((bid) => {
    const latest = costRows.find((cost) => cost.plan === bid.plan && cost.code === bid.code);
    return latest ? { ...bid, current: latest.current, description: latest.description, trade: latest.trade } : bid;
  }).filter((row) => {
    const matchesFilters = (plan === 'All plans' || row.plan === plan) && (trade === 'All trades' || row.trade === trade);
    return matchesFilters && `${row.plan} ${row.trade} ${row.code} ${row.description} ${row.bidder}`.toLowerCase().includes(query.toLowerCase());
  });
  const bidderSummary = useMemo(() => {
    const grouped = new Map<string, { count: number; variance: number; current: number }>();
    visibleBids.forEach((row) => {
      const current = grouped.get(row.bidder) ?? { count: 0, variance: 0, current: 0 };
      current.count += 1; current.variance += row.amount - row.current; current.current += row.current; grouped.set(row.bidder, current);
    });
    return [...grouped.entries()].map(([name, value]) => ({ name, ...value, percent: value.current ? (value.variance / value.current) * 100 : 0 }));
  }, [visibleBids]);

  async function loadBids() {
    try {
      const response = await fetch('/api/bids', { cache: 'no-store' });
      const result = (await response.json()) as { bids?: Bid[]; message?: string };
      if (!response.ok) throw new Error(result.message ?? 'Could not load shared bids.');
      setBids(result.bids ?? []);
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'Could not load shared bids.');
    }
  }

  async function saveBid(event: FormEvent) {
    event.preventDefault();
    const numericAmount = Number(amount.replace(/[^0-9.-]/g, ''));
    if (!selected) { setNotice('The Floorplans cost data must be available before a bid can be saved.'); return; }
    if (!bidder.trim() || !Number.isFinite(numericAmount)) { setNotice('Add a bidder and a valid amount.'); return; }
    const company = bidder.trim();
    setSaving(true); setNotice('');
    try {
      const response = await fetch('/api/bids', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ...selected, bidder: company, amount: numericAmount, date: bidDate, scope, notes }),
      });
      const result = (await response.json()) as { bid?: Bid; message?: string };
      if (!response.ok || !result.bid) throw new Error(result.message ?? 'Could not save the shared bid.');
      setBids((current) => [result.bid!, ...current]);
      setBidder(''); setAmount(''); setNotes(''); setNotice(`Saved ${company}'s bid against ${selected.code} for everyone.`);
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'Could not save the shared bid. Your entry is still in the form.');
    } finally { setSaving(false); }
  }

  async function deleteBid(bidId: string) {
    try {
      const response = await fetch(`/api/bids?id=${encodeURIComponent(bidId)}`, { method: 'DELETE' });
      const result = (await response.json()) as { message?: string };
      if (!response.ok) throw new Error(result.message ?? 'Could not delete the bid.');
      setBids((current) => current.filter((bid) => bid.bidId !== bidId));
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'Could not delete the bid.');
    }
  }

  return (
    <main className="min-h-screen bg-background text-foreground">
      <header className="border-b border-border bg-[#080b0d]">
        <div className="mx-auto flex max-w-[1500px] flex-col gap-5 px-5 py-6 lg:flex-row lg:items-end lg:justify-between lg:px-8">
          <div><p className="flex items-center gap-2 text-xs font-bold uppercase tracking-[0.18em] text-primary"><Building2 size={15} /> Builder Bid Desk</p><h1 className="mt-2 text-3xl font-semibold tracking-[-0.035em] text-white">Trade Cost Dashboard</h1><p className="mt-2 text-sm text-slate-400">{plans.length} plans · {trades.length} trades · {costRows.length} current cost lines · {bids.length} bids</p></div>
          <div className="grid w-full gap-3 sm:grid-cols-2 lg:max-w-[720px]">
            <label className="field-label">Plan<select value={plan} onChange={(e) => setPlan(e.target.value)}><option>All plans</option>{plans.map((value) => <option key={value}>{value}</option>)}</select></label>
            <label className="field-label">Trade<select value={trade} onChange={(e) => setTrade(e.target.value)}><option>All trades</option>{trades.map((value) => <option key={value}>{value}</option>)}</select></label>
          </div>
        </div>
      </header>
      <section className="mx-auto grid max-w-[1500px] gap-5 px-5 py-5 lg:grid-cols-[360px_minmax(0,1fr)] lg:px-8">
        <aside className="space-y-5">
          <form onSubmit={saveBid} className="panel p-5">
            <div className="mb-5 flex items-center justify-between"><h2>New bid</h2><span className="badge"><Plus size={13} /> Fast entry</span></div>
            <div className="space-y-3">
              <label className="field-label">Plan<select value={bidPlan} onChange={(e) => { const nextPlan = e.target.value; setBidPlan(nextPlan); const match = costRows.find((row) => row.plan === nextPlan); if (match) setSelectedId(match.id); }}>{plans.map((value) => <option key={value}>{value}</option>)}</select></label>
              <label className="field-label">Trade<select value={selected?.id ?? ''} onChange={(e) => setSelectedId(Number(e.target.value))} disabled={!selected}>{bidCostOptions.map((row) => <option value={row.id} key={row.id}>{row.description}</option>)}</select></label>
              <label className="field-label">Bidder<input value={bidder} onChange={(e) => setBidder(e.target.value)} placeholder="Trade company" /></label>
              <div className="grid grid-cols-2 gap-3"><label className="field-label">Amount<input value={amount} onChange={(e) => setAmount(e.target.value)} inputMode="decimal" placeholder="$0" /></label><label className="field-label">Date<input type="date" value={bidDate} onChange={(e) => setBidDate(e.target.value)} /></label></div>
              <label className="field-label">Scope<input value={scope} onChange={(e) => setScope(e.target.value)} placeholder="Base bid" /></label>
              <label className="field-label">Notes<textarea value={notes} onChange={(e) => setNotes(e.target.value)} placeholder="Allowances, exclusions, alternates" rows={3} /></label>
              <button className="primary-button" type="submit" disabled={saving || !selected}><Check size={16} /> {saving ? 'Saving…' : 'Save bid'}</button>
            </div>
          </form>
          <div className="panel p-5"><div className="flex items-start justify-between gap-4"><div><p className="eyebrow">Company data</p><h2 className="mt-1">SharePoint source</h2></div><span className="status-dot">{costStatus}</span></div><div className="mt-4 rounded-xl border border-border bg-black/15 p-4"><div className="flex items-center gap-3"><FileSpreadsheet className="text-primary" size={21} /><div><p className="text-sm font-semibold">Master Costing Sheet</p><p className="mt-0.5 text-xs text-muted-foreground">{costRows.length.toLocaleString()} cost lines · {plans.length} plans</p><p className="mt-1 text-[11px] text-muted-foreground">Imported {sourceData.source.importedAt}</p></div></div><button className="secondary-button mt-4" type="button" onClick={() => window.open(refreshWorkflowUrl, '_blank', 'noopener,noreferrer')}><RefreshCw size={15} />Refresh via GitHub</button></div></div>
        </aside>
        <div className="min-w-0 space-y-5">
          <section className="panel p-5"><div className="mb-4 flex flex-wrap items-end justify-between gap-3"><div><h2>Bidder summary</h2><p className="mt-1 text-sm text-muted-foreground">{bidderSummary.length} bidders · {visibleBids.length} active bids</p></div><div className="cloud-pill"><Cloud size={14} /> Latest cost baseline</div></div><div className="overflow-x-auto"><table><thead><tr><th>Bidder</th><th className="text-right">Active bids</th><th className="text-right">Avg. difference</th><th className="text-right">Avg. %</th></tr></thead><tbody>{bidderSummary.map((row) => <tr key={row.name}><td className="font-semibold text-white">{row.name}</td><td className="text-right">{row.count}</td><td className={`text-right font-semibold ${row.variance <= 0 ? 'good' : 'bad'}`}>{row.variance <= 0 ? '' : '+'}{money.format(row.variance / row.count)}</td><td className={`text-right font-semibold ${row.percent <= 0 ? 'good' : 'bad'}`}>{row.percent > 0 ? '+' : ''}{row.percent.toFixed(1)}%</td></tr>)}</tbody></table></div></section>
          <section className="panel overflow-hidden"><div className="flex flex-col gap-4 border-b border-border p-5 sm:flex-row sm:items-center sm:justify-between"><div><h2>Cost comparison</h2><p className="mt-1 text-sm text-muted-foreground">Current cost vs. potential trade bid</p></div><label className="search"><Search size={16} /><input aria-label="Search bids" value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search plan, code, bidder…" /></label></div><div className="overflow-x-auto"><table><thead><tr><th>Bid date</th><th>Plan / cost code</th><th>Trade</th><th>Bidder</th><th className="text-right">Current</th><th className="text-right">Bid</th><th className="text-right">Variance</th><th className="text-right">%</th><th /></tr></thead><tbody>
            {visibleBids.map((row) => { const variance = row.amount - row.current; const percent = row.current ? (variance / row.current) * 100 : 0; const favorable = variance <= 0; return <tr key={row.bidId}><td className="whitespace-nowrap text-slate-400">{dateLabel.format(new Date(`${row.date}T12:00:00`))}</td><td><strong className="text-white">{row.plan}</strong><small>{row.code}</small></td><td><strong className="text-white">{row.description}</strong><small>{row.trade}</small></td><td>{row.bidder}</td><td className="text-right font-semibold text-white">{money.format(row.current)}</td><td className="text-right font-semibold text-white">{money.format(row.amount)}</td><td className={`text-right font-semibold ${favorable ? 'good' : 'bad'}`}>{favorable ? <ArrowDownRight className="inline" size={14} /> : <ArrowUpRight className="inline" size={14} />}{money.format(Math.abs(variance))}</td><td className={`text-right font-semibold ${favorable ? 'good' : 'bad'}`}>{percent > 0 ? '+' : ''}{percent.toFixed(1)}%</td><td><button className="icon-button" aria-label={`Delete ${row.bidder} bid`} onClick={() => void deleteBid(row.bidId)}><Trash2 size={15} /></button></td></tr>; })}
            {!visibleBids.length && <tr><td colSpan={9} className="py-14 text-center text-muted-foreground">No bids match these filters.</td></tr>}
          </tbody></table></div></section>
        </div>
      </section>
      {notice && <output className="toast" aria-live="polite">{notice}<button onClick={() => setNotice('')} aria-label="Dismiss">×</button></output>}
    </main>
  );
}
