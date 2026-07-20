import { NextResponse } from 'next/server';
import { SYNC_SOURCES, fetchSheet, fetchReport, processSource, normalizeMonth, KpiRow } from '@/lib/smartsheet';

export const dynamic = 'force-dynamic';

// Smartsheet webhook verification (challenge-response)
export async function POST(req: Request) {
    const body = await req.json();

  // Smartsheet sends a verification request with a challenge
  if (body.challenge) {
        return NextResponse.json(
          { smartsheetHookResponse: body.challenge },
          { status: 200, headers: { 'Content-Type': 'application/json' } }
              );
  }

  // Webhook event received — trigger a full sync
  const token = process.env.SMARTSHEET_TOKEN;
    if (!token) return NextResponse.json({ error: 'SMARTSHEET_TOKEN not set' }, { status: 500 });

  const results: Record<string, { rows: KpiRow[]; error?: string }> = {};
    const errors: string[] = [];

  await Promise.allSettled(
        SYNC_SOURCES.map(async (src) => {
                try {
                          const raw = src.sheetId
                            ? await fetchSheet(src.sheetId, token)
                                      : await fetchReport(src.reportId!, token);
                          const processed = processSource(src, raw);
                          results[src.key] = { rows: processed };
                } catch (e: any) {
                          errors.push(`${src.key}: ${e.message}`);
                          results[src.key] = { rows: [], error: e.message };
                }
        })
      );

  return NextResponse.json({
        ok: true,
        source: 'webhook',
        syncedAt: new Date().toISOString(),
        keys: Object.keys(results),
        errors
  });
}

// GET for health check
export async function GET() {
    return NextResponse.json({ status: 'webhook endpoint ready' });
}
