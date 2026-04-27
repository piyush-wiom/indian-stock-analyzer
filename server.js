'use strict';

const fs = require('fs');
const credPath = 'C:\\credentials\\.env';
if (fs.existsSync(credPath)) {
  require('dotenv').config({ path: credPath });
} else {
  require('dotenv').config();
}

const express  = require('express');
const path     = require('path');
const multer   = require('multer');
const ExcelJS  = require('exceljs');
const https    = require('https');

const YahooFinanceClass = require('yahoo-finance2').default;
const yf = new YahooFinanceClass({ suppressNotices: ['ripHistorical', 'yahooSurvey'] });

const { RSI, MACD, BollingerBands, SMA } = require('technicalindicators');

const app    = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 5 * 1024 * 1024 } });

app.use(express.json({ limit: '2mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ─────────────────────────────────────────────
//  Ticker resolution
// ─────────────────────────────────────────────
const ALIASES = {
  'tcs':'TCS.NS','tata consultancy':'TCS.NS',
  'infosys':'INFY.NS','infy':'INFY.NS',
  'wipro':'WIPRO.NS',
  'hcl tech':'HCLTECH.NS','hcltech':'HCLTECH.NS',
  'tech mahindra':'TECHM.NS','techm':'TECHM.NS',
  'reliance':'RELIANCE.NS','reliance industries':'RELIANCE.NS',
  'hdfc bank':'HDFCBANK.NS','hdfcbank':'HDFCBANK.NS','hdfc':'HDFCBANK.NS',
  'icici bank':'ICICIBANK.NS','icicibank':'ICICIBANK.NS','icici':'ICICIBANK.NS',
  'sbi':'SBIN.NS','state bank':'SBIN.NS','sbin':'SBIN.NS',
  'axis bank':'AXISBANK.NS','axisbank':'AXISBANK.NS',
  'kotak bank':'KOTAKBANK.NS','kotak':'KOTAKBANK.NS','kotakbank':'KOTAKBANK.NS',
  'bajaj finance':'BAJFINANCE.NS','bajfinance':'BAJFINANCE.NS',
  'bajaj finserv':'BAJAJFINSV.NS','bajajfinsv':'BAJAJFINSV.NS',
  'bajaj auto':'BAJAJ-AUTO.NS',
  'maruti':'MARUTI.NS','maruti suzuki':'MARUTI.NS',
  'tata motors':'TATAMOTORS.NS','tatamotors':'TATAMOTORS.NS',
  'tata steel':'TATASTEEL.NS','tatasteel':'TATASTEEL.NS',
  'mahindra':'M&M.NS','m&m':'M&M.NS',
  'sun pharma':'SUNPHARMA.NS','sunpharma':'SUNPHARMA.NS',
  'dr reddy':'DRREDDY.NS','drreddy':'DRREDDY.NS',
  'cipla':'CIPLA.NS',
  'l&t':'LT.NS','larsen':'LT.NS',
  'ntpc':'NTPC.NS','ongc':'ONGC.NS',
  'coal india':'COALINDIA.NS','coalindia':'COALINDIA.NS',
  'airtel':'BHARTIARTL.NS','bharti airtel':'BHARTIARTL.NS','bhartiartl':'BHARTIARTL.NS',
  'itc':'ITC.NS',
  'hul':'HINDUNILVR.NS','hindustan unilever':'HINDUNILVR.NS',
  'nestle':'NESTLEIND.NS','nestle india':'NESTLEIND.NS',
  'adani ports':'ADANIPORTS.NS','adaniports':'ADANIPORTS.NS',
  'adani enterprises':'ADANIENT.NS','adanient':'ADANIENT.NS',
  'power grid':'POWERGRID.NS','powergrid':'POWERGRID.NS',
  'indusindbk':'INDUSINDBK.NS','indusind':'INDUSINDBK.NS',
  'yes bank':'YESBANK.NS','yesbank':'YESBANK.NS',
  'ltim':'LTIM.NS','ltimindtree':'LTIM.NS',
  'persistent':'PERSISTENT.NS',
  'mphasis':'MPHASIS.NS',
  'hero motocorp':'HEROMOTOCO.NS','heromotoco':'HEROMOTOCO.NS',
  'eicher motors':'EICHERMOT.NS','eichermot':'EICHERMOT.NS',
  'tvs motor':'TVSMOTOR.NS','tvsmotor':'TVSMOTOR.NS',
  'britannia':'BRITANNIA.NS',
  'dabur':'DABUR.NS','marico':'MARICO.NS',
  'bpcl':'BPCL.NS','ioc':'IOC.NS','gail':'GAIL.NS',
  'lic':'LICI.NS','lici':'LICI.NS',
  'hdfc life':'HDFCLIFE.NS','hdfclife':'HDFCLIFE.NS',
  'sbi life':'SBILIFE.NS','sbilife':'SBILIFE.NS',
  'nifty':'^NSEI','banknifty':'^NSEBANK','bank nifty':'^NSEBANK','sensex':'^BSESN',
};

function resolveTicker(input) {
  const raw   = (input || '').trim();
  const lower = raw.toLowerCase();
  if (ALIASES[lower]) return ALIASES[lower];
  for (const [k, v] of Object.entries(ALIASES)) {
    if (lower.includes(k) || k.includes(lower)) return v;
  }
  if (/^\d+$/.test(raw)) return raw + '.BO';
  if (raw.toUpperCase().endsWith('.NS') || raw.toUpperCase().endsWith('.BO')) return raw.toUpperCase();
  return raw.toUpperCase() + '.NS';
}

// ─────────────────────────────────────────────
//  Technical Analysis
// ─────────────────────────────────────────────
function analyse(closes, volumes) {
  const n = closes.length;

  // RSI
  let rsiVal = 50;
  try { const a = RSI.calculate({ values: closes, period: 14 }); rsiVal = a[a.length - 1] ?? 50; } catch (_) {}

  // MACD
  let macdVal = 0, signalVal = 0, histVal = 0;
  try {
    const a = MACD.calculate({ values: closes, fastPeriod: 12, slowPeriod: 26, signalPeriod: 9, SimpleMAOscillator: false, SimpleMASignal: false });
    const last = a[a.length - 1];
    if (last) { macdVal = last.MACD || 0; signalVal = last.signal || 0; histVal = last.histogram || 0; }
  } catch (_) {}

  // Bollinger
  let bbU = closes[n - 1], bbM = closes[n - 1], bbL = closes[n - 1];
  try {
    const a = BollingerBands.calculate({ values: closes, period: 20, stdDev: 2 });
    const last = a[a.length - 1];
    if (last) { bbU = last.upper; bbM = last.middle; bbL = last.lower; }
  } catch (_) {}

  // SMA
  let sma50 = null, sma200 = null;
  try {
    if (n >= 50)  { const a = SMA.calculate({ values: closes, period: 50  }); sma50  = a[a.length - 1]; }
    if (n >= 200) { const a = SMA.calculate({ values: closes, period: 200 }); sma200 = a[a.length - 1]; }
  } catch (_) {}

  // Volume trend
  const vol5  = volumes.slice(-5).reduce((a, b) => a + b, 0) / 5;
  const vol20 = volumes.slice(-20).reduce((a, b) => a + b, 0) / 20;

  // Scores
  const rsiScore = rsiVal < 30 ? 2 : rsiVal < 45 ? 1 : rsiVal > 70 ? -2 : rsiVal > 55 ? -1 : 0;
  const macdScore = macdVal > signalVal ? 1 : -1;
  const maScore = (sma50 && sma200) ? (sma50 > sma200 ? 1 : -1) : 0;
  const cp = closes[n - 1];
  const bbScore = cp < bbL ? 1 : cp > bbU ? -1 : 0;
  const volScore = vol20 > 0 && vol5 > vol20 * 1.2 ? 0.5 : 0;
  const total = rsiScore + macdScore + maScore + bbScore + volScore;

  let rec, recDetail, recForNonHolder, recForHolder, color, emoji;
  if      (total >= 3)  {
    rec = 'Strong Buy';          recForNonHolder = 'Strong Buy';        recForHolder = 'Add More';
    recDetail = 'Strong entry opportunity. Technicals are aligned bullishly.';
    color = '#00A36C'; emoji = '🚀';
  } else if (total >= 1)  {
    rec = 'Buy';                 recForNonHolder = 'Good Entry';         recForHolder = 'Hold & Add';
    recDetail = 'Good time to enter. Consider buying in parts (SIP style).';
    color = '#2ECC71'; emoji = '✅';
  } else if (total >= -1) {
    rec = 'Wait & Watch';        recForNonHolder = 'Wait for Entry';     recForHolder = 'Hold — No Action';
    recDetail = 'No clear signal. If you own it — hold. If not — wait for a better entry.';
    color = '#F39C12'; emoji = '⏸️';
  } else if (total >= -3) {
    rec = 'Avoid / Cut Loss';    recForNonHolder = 'Avoid Buying';       recForHolder = 'Review Your Position';
    recDetail = 'Technicals are weak. If you own it — consider exiting. If not — avoid buying now.';
    color = '#E74C3C'; emoji = '⚠️';
  } else {
    rec = 'Strong Avoid / Exit'; recForNonHolder = 'Stay Away';          recForHolder = 'Consider Exiting';
    recDetail = 'High risk. Strong downtrend. If you own it — exit. If not — stay away.';
    color = '#8B0000'; emoji = '🔴';
  }

  const confidence = Math.min(100, Math.round(Math.abs(total) / 5.5 * 100));

  return {
    rsi:            { value: +rsiVal.toFixed(2), score: rsiScore },
    macd:           { macd: +macdVal.toFixed(4), signal: +signalVal.toFixed(4), histogram: +histVal.toFixed(4), score: macdScore },
    movingAverages: { sma50: sma50 ? +sma50.toFixed(2) : null, sma200: sma200 ? +sma200.toFixed(2) : null, score: maScore, crossType: (sma50 && sma200) ? (sma50 > sma200 ? 'golden' : 'death') : 'neutral' },
    bollinger:      { upper: +bbU.toFixed(2), middle: +bbM.toFixed(2), lower: +bbL.toFixed(2), score: bbScore },
    volume:         { avg5d: Math.round(vol5), avg20d: Math.round(vol20), score: volScore },
    totalScore:     +total.toFixed(2),
    recommendation: rec, recForNonHolder, recForHolder, recDetail, color, emoji, confidence,
  };
}

function buildExplanation(ind, name) {
  const pos = [], neg = [];
  if (ind.rsi.score >= 1)               pos.push('RSI near oversold zone — selling overdone');
  else if (ind.rsi.score <= -1)         neg.push('RSI in overbought zone — buying pressure high');
  if (ind.macd.score === 1)             pos.push('MACD bullish crossover — upward momentum');
  else                                  neg.push('MACD bearish signal — downward momentum');
  if (ind.movingAverages.score === 1)   pos.push('Golden Cross active — long-term uptrend');
  else if (ind.movingAverages.score === -1) neg.push('Death Cross active — long-term downtrend');
  if (ind.bollinger.score === 1)        pos.push('Price below lower Bollinger Band — oversold bounce possible');
  else if (ind.bollinger.score === -1)  neg.push('Price above upper Bollinger Band — overbought');
  if (ind.volume.score === 0.5)         pos.push('Volume spike — signal strength confirmed');

  let t = `${name}: ${ind.recommendation} — ${ind.confidence}% confidence.\n📌 ${ind.recDetail}`;
  if (pos.length) t += '\n\n✅ Bullish signals: ' + pos.join(', ') + '.';
  if (neg.length) t += '\n⚠️ Bearish signals: ' + neg.join(', ') + '.';
  t += '\n\n⚡ Disclaimer: Educational only. Consult a SEBI-registered advisor before investing.';
  return t;
}

// ─────────────────────────────────────────────
//  Portfolio Decision Engine
// ─────────────────────────────────────────────
function portfolioDecision(techScore, pnlPct, ind, currentPrice, buyPrice, week52High, quantity) {
  const sma200    = ind.movingAverages.sma200;
  const crossType = ind.movingAverages.crossType;
  const rsi       = ind.rsi.value;

  // ── Small position guard ─────────────────────
  // Don't recommend partial exits for tiny positions (< 15 shares or < ₹50,000 value)
  const positionValue = quantity ? currentPrice * quantity : Infinity;
  const isSmallPosition = quantity && (quantity < 15 || positionValue < 50000);

  // ── RSI Recovery check ───────────────────────
  // If RSI is above 50 and rising, stock is recovering — soften bearish calls
  const isRecovering = rsi > 50;

  // ── Computed values ──────────────────────────
  // Stop loss MUST always be below current price.
  // In profit  → trail from buy price OR 7% below current (whichever is higher — protects gains)
  // In loss    → 7% below current price only (buy price is already above current — irrelevant)
  let stopLoss;
  if (currentPrice >= buyPrice) {
    stopLoss = +Math.max(buyPrice, currentPrice * 0.93).toFixed(2);
  } else {
    stopLoss = +(currentPrice * 0.93).toFixed(2);
  }

  let targetPrice;
  if (sma200 && currentPrice < sma200) {
    targetPrice = +sma200.toFixed(2);
  } else {
    targetPrice = +(currentPrice * 1.15).toFixed(2);
  }
  if (week52High && targetPrice > week52High) targetPrice = +week52High.toFixed(2);

  const breakEvenGainPct = buyPrice > currentPrice
    ? +((buyPrice - currentPrice) / currentPrice * 100).toFixed(2)
    : 0;
  const rr = (currentPrice - stopLoss) > 0
    ? +((targetPrice - currentPrice) / (currentPrice - stopLoss)).toFixed(2)
    : 0;

  // ── Override rules ───────────────────────────
  // Never average into Death Cross when already in loss
  const deathCrossBlock = (crossType === 'death' && pnlPct < -5);
  // Boost averaging signal on Golden Cross
  const goldenBoost = (crossType === 'golden' && pnlPct < 0) ? 0.5 : 0;
  const effectiveScore = techScore + goldenBoost;

  // Overbought + in profit → force book profit
  if (rsi > 75 && pnlPct > 10) {
    return {
      action: 'BOOK PROFIT', urgency: 'HIGH', actionColor: '#C0392B',
      urgencyBadge: '🟠 HIGH',
      timeHorizon: 'Exit within 1 week — RSI overbought while in profit',
      reasoning: [
        `RSI at ${rsi.toFixed(1)} — stock is significantly overbought`,
        `You are already up ${pnlPct.toFixed(1)}% from your buy price`,
        'Overbought conditions + existing profit = ideal exit window',
        'Re-enter on pullback when RSI cools below 55',
      ],
      stopLoss, targetPrice, averageAt: null, breakEvenGainPct, riskRewardRatio: rr,
    };
  }

  // ── Severe capital loss — exit regardless ────
  if (pnlPct <= -40) {
    return {
      action: 'CUT LOSS', urgency: 'CRITICAL', actionColor: '#6B0000',
      urgencyBadge: '🔴 CRITICAL',
      timeHorizon: 'Exit within 1–3 trading days',
      reasoning: [
        `Stock is down ${Math.abs(pnlPct).toFixed(1)}% from your buy price — severe capital destruction`,
        'At this level, averaging down would require even bigger recovery to break even',
        `You need ${breakEvenGainPct.toFixed(1)}% gain just to break even — very difficult`,
        'Preserve remaining capital; re-evaluate after 3–6 months stabilisation',
      ],
      stopLoss, targetPrice, averageAt: null, breakEvenGainPct, riskRewardRatio: rr,
    };
  }

  // ── Strong Sell zone (score < -3) ────────────
  if (techScore < -3) {
    if (pnlPct < 0) {
      return {
        action: 'EXIT NOW', urgency: 'CRITICAL', actionColor: '#6B0000',
        urgencyBadge: '🔴 CRITICAL',
        timeHorizon: 'Exit immediately',
        reasoning: [
          `All 5 technical indicators are bearish (score ${techScore.toFixed(1)})`,
          `You are already at a ${Math.abs(pnlPct).toFixed(1)}% loss`,
          'Continuing to hold risks further significant decline',
          'Stop loss breached — exit to cut further damage',
        ],
        stopLoss, targetPrice, averageAt: null, breakEvenGainPct, riskRewardRatio: rr,
      };
    } else {
      return {
        action: 'BOOK ALL PROFIT', urgency: 'HIGH', actionColor: '#C0392B',
        urgencyBadge: '🟠 HIGH',
        timeHorizon: 'Exit within 1 week before reversal',
        reasoning: [
          `Strong Sell signal (score ${techScore.toFixed(1)}) — momentum reversing sharply`,
          `You are up ${pnlPct.toFixed(1)}% — lock in the gains now`,
          'RSI, MACD and moving averages all pointing downward',
          'Preserve profits; look for re-entry after correction',
        ],
        stopLoss, targetPrice, averageAt: null, breakEvenGainPct, riskRewardRatio: rr,
      };
    }
  }

  // ── Sell zone (-3 to -1) ─────────────────────
  if (techScore < -1) {
    if (pnlPct < -10) {
      // If price is recovering (RSI > 50), soften CUT LOSS → HOLD with tight SL
      if (isRecovering) {
        return {
          action: 'HOLD — TIGHT SL', urgency: 'MEDIUM', actionColor: '#D68910',
          urgencyBadge: '🟡 MEDIUM',
          timeHorizon: 'Hold but protect with stop loss',
          reasoning: [
            `Stock is down ${Math.abs(pnlPct).toFixed(1)}% but RSI ${rsi.toFixed(0)} shows price recovering`,
            'Technicals are mixed — short-term bearish but momentum improving',
            `Keep strict stop loss at ₹${stopLoss} — exit if this level breaks`,
            'If RSI crosses 60 and price holds, technicals will improve',
          ],
          stopLoss, targetPrice, averageAt: null, breakEvenGainPct, riskRewardRatio: rr,
        };
      }
      return {
        action: 'CUT LOSS', urgency: 'HIGH', actionColor: '#C0392B',
        urgencyBadge: '🟠 HIGH',
        timeHorizon: 'Exit within 1 week',
        reasoning: [
          `Bearish technical signal (score ${techScore.toFixed(1)}) combined with ${Math.abs(pnlPct).toFixed(1)}% loss`,
          'Risk of further decline is high — technicals confirm weakness',
          deathCrossBlock ? 'Death Cross active — long-term trend is negative' : `RSI at ${rsi.toFixed(0)} — no recovery signal yet`,
          'Cutting loss now limits damage vs holding in deteriorating stock',
        ],
        stopLoss, targetPrice, averageAt: null, breakEvenGainPct, riskRewardRatio: rr,
      };
    } else if (pnlPct < 0) {
      return {
        action: 'CUT SMALL LOSS', urgency: 'MEDIUM', actionColor: '#D35400',
        urgencyBadge: '🟡 MEDIUM',
        timeHorizon: 'Exit within 2 weeks',
        reasoning: [
          `Bearish signal (score ${techScore.toFixed(1)}) with a small ${Math.abs(pnlPct).toFixed(1)}% loss`,
          'Better to exit with a small loss than wait for it to deepen',
          'MACD and RSI suggesting downside pressure ahead',
          `Re-enter if technicals improve and price holds above ₹${stopLoss}`,
        ],
        stopLoss, targetPrice, averageAt: null, breakEvenGainPct, riskRewardRatio: rr,
      };
    } else if (pnlPct > 15) {
      // Small position → don't tell them to sell, just protect with SL
      if (isSmallPosition) {
        return {
          action: 'HOLD — PROTECT GAINS', urgency: 'LOW', actionColor: '#1E8449',
          urgencyBadge: '🟢 LOW',
          timeHorizon: 'Hold — position too small to partially exit',
          reasoning: [
            `Small position (${quantity} shares, ₹${Math.round(positionValue).toLocaleString('en-IN')}) — partial exit not practical`,
            `You are up ${pnlPct.toFixed(1)}% — let it ride with a trailing stop`,
            `Protect gains: move stop loss up to ₹${stopLoss}`,
            'Exit fully only if stop loss is breached',
          ],
          stopLoss, targetPrice, averageAt: null, breakEvenGainPct, riskRewardRatio: rr,
        };
      }
      return {
        action: 'BOOK PROFIT', urgency: 'HIGH', actionColor: '#C0392B',
        urgencyBadge: '🟠 HIGH',
        timeHorizon: 'Exit within 1 week',
        reasoning: [
          `Sell signal (score ${techScore.toFixed(1)}) while you are up ${pnlPct.toFixed(1)}%`,
          'Momentum is weakening — protect your gains before reversal',
          'MACD turning bearish and RSI losing strength',
          `Target ₹${targetPrice} may not be reached in near term`,
        ],
        stopLoss, targetPrice, averageAt: null, breakEvenGainPct, riskRewardRatio: rr,
      };
    } else {
      // Small position in small profit → just hold
      if (isSmallPosition) {
        return {
          action: 'HOLD', urgency: 'LOW', actionColor: '#888',
          urgencyBadge: '⚪ LOW',
          timeHorizon: 'Hold — position too small to partially exit',
          reasoning: [
            `Small position (${quantity} shares) — partial exit not practical`,
            `Technicals slightly bearish (score ${techScore.toFixed(1)}) but not enough to exit entirely`,
            `Keep stop loss at ₹${stopLoss} — exit fully only if breached`,
            'Re-evaluate when position size is larger',
          ],
          stopLoss, targetPrice, averageAt: null, breakEvenGainPct, riskRewardRatio: rr,
        };
      }
      return {
        action: 'BOOK PROFIT', urgency: 'MEDIUM', actionColor: '#D35400',
        urgencyBadge: '🟡 MEDIUM',
        timeHorizon: 'Exit within 2 weeks',
        reasoning: [
          `Bearish technical signal (score ${techScore.toFixed(1)})`,
          `You are up ${pnlPct.toFixed(1)}% — book while you can`,
          'Technicals suggest more downside than upside ahead',
          'Wait for a new buy signal before re-entering',
        ],
        stopLoss, targetPrice, averageAt: null, breakEvenGainPct, riskRewardRatio: rr,
      };
    }
  }

  // ── Strong Buy zone (score ≥ 3) ──────────────
  if (effectiveScore >= 3) {
    if (pnlPct < -25) {
      return {
        action: 'HOLD — DO NOT AVERAGE', urgency: 'MEDIUM', actionColor: '#D68910',
        urgencyBadge: '🟡 MEDIUM',
        timeHorizon: '3–6 months — wait for price stabilisation before adding',
        reasoning: [
          `Strong technical signal (score ${effectiveScore.toFixed(1)}) but ${Math.abs(pnlPct).toFixed(1)}% loss is very deep`,
          `Need ${breakEvenGainPct.toFixed(1)}% gain to break even — high recovery required`,
          'Averaging now risks deploying more capital in a heavily impaired position',
          'Wait for price to stabilise above SMA50 before adding more',
        ],
        stopLoss, targetPrice, averageAt: null, breakEvenGainPct, riskRewardRatio: rr,
      };
    } else if (pnlPct < -5) {
      const avgLadder = [
        +currentPrice.toFixed(2),
        +(currentPrice * 0.96).toFixed(2),
        +(currentPrice * 0.92).toFixed(2),
      ];
      return {
        action: 'STRONG AVERAGE', urgency: 'OPPORTUNITY', actionColor: '#1A8A5A',
        urgencyBadge: '🟢 OPPORTUNITY',
        timeHorizon: '3–6 months — DCA in 3 tranches',
        reasoning: [
          `Strong Buy signal (score ${effectiveScore.toFixed(1)}) while stock is at a ${Math.abs(pnlPct).toFixed(1)}% discount from your cost`,
          crossType === 'golden' ? 'Golden Cross confirms long-term uptrend — ideal averaging window' : 'RSI near oversold + MACD turning bullish',
          `DCA ladder: buy at ₹${avgLadder[0]} now, more at ₹${avgLadder[1]}, more at ₹${avgLadder[2]}`,
          `Target ₹${targetPrice} — Risk:Reward = ${rr}:1`,
        ],
        stopLoss, targetPrice, averageAt: avgLadder, breakEvenGainPct, riskRewardRatio: rr,
      };
    } else {
      return {
        action: 'HOLD & ADD', urgency: 'LOW', actionColor: '#1E8449',
        urgencyBadge: '🟢 LOW',
        timeHorizon: '6+ months — riding a winner',
        reasoning: [
          `Strong Buy signal (score ${effectiveScore.toFixed(1)}) with a profitable position`,
          pnlPct > 0 ? `You are up ${pnlPct.toFixed(1)}% — momentum in your favour` : 'Near breakeven with strong technicals',
          'Hold core position; add more on any 3–5% dips',
          `Target ₹${targetPrice} | Stop Loss ₹${stopLoss}`,
        ],
        stopLoss, targetPrice, averageAt: [+(currentPrice * 0.97).toFixed(2)], breakEvenGainPct, riskRewardRatio: rr,
      };
    }
  }

  // ── Buy zone (1 to 3) ────────────────────────
  if (effectiveScore >= 1) {
    if (pnlPct < -20) {
      return {
        action: 'HOLD', urgency: 'MEDIUM', actionColor: '#D68910',
        urgencyBadge: '🟡 MEDIUM',
        timeHorizon: '2–4 months',
        reasoning: [
          `Mild Buy signal (score ${effectiveScore.toFixed(1)}) but position is down ${Math.abs(pnlPct).toFixed(1)}%`,
          'Signal not strong enough to justify averaging into a large loss',
          deathCrossBlock ? 'Death Cross still active — avoid adding until cross recovers' : 'Hold and wait for technical confirmation to strengthen',
          `If price falls below ₹${stopLoss}, consider exiting`,
        ],
        stopLoss, targetPrice, averageAt: null, breakEvenGainPct, riskRewardRatio: rr,
      };
    } else if (pnlPct < -5 && !deathCrossBlock) {
      const avgLadder = [
        +currentPrice.toFixed(2),
        +(currentPrice * 0.97).toFixed(2),
      ];
      return {
        action: 'AVERAGE', urgency: 'OPPORTUNITY', actionColor: '#1A6B9A',
        urgencyBadge: '🔵 OPPORTUNITY',
        timeHorizon: '2–4 months',
        reasoning: [
          `Buy signal (score ${effectiveScore.toFixed(1)}) with position at ${Math.abs(pnlPct).toFixed(1)}% discount`,
          'Moderate averaging opportunity — buy in 2 tranches',
          `Average at ₹${avgLadder[0]} and again at ₹${avgLadder[1]} on dips`,
          `Target ₹${targetPrice} | Stop Loss ₹${stopLoss}`,
        ],
        stopLoss, targetPrice, averageAt: avgLadder, breakEvenGainPct, riskRewardRatio: rr,
      };
    } else if (pnlPct > 30) {
      return {
        action: 'PARTIAL PROFIT', urgency: 'MEDIUM', actionColor: '#D35400',
        urgencyBadge: '🟡 MEDIUM',
        timeHorizon: 'Book 50% within 2 weeks, hold rest',
        reasoning: [
          `Up ${pnlPct.toFixed(1)}% — significant gains on the table`,
          `Buy signal (score ${effectiveScore.toFixed(1)}) still holds so do not exit fully`,
          'Book 50% profits to lock in gains; hold rest for further upside',
          `Raise stop loss on remaining position to ₹${stopLoss}`,
        ],
        stopLoss, targetPrice, averageAt: null, breakEvenGainPct, riskRewardRatio: rr,
      };
    } else {
      return {
        action: 'HOLD', urgency: 'LOW', actionColor: '#5D6D7E',
        urgencyBadge: '🟢 LOW',
        timeHorizon: '1–3 months',
        reasoning: [
          `Buy signal (score ${effectiveScore.toFixed(1)}) — stock is on positive trajectory`,
          pnlPct >= 0 ? `You are up ${pnlPct.toFixed(1)}% — comfortable position` : `Minor ${Math.abs(pnlPct).toFixed(1)}% loss, recovery expected`,
          `Target ₹${targetPrice} | Stop Loss ₹${stopLoss}`,
          'Review again in 4–6 weeks or on major news',
        ],
        stopLoss, targetPrice, averageAt: null, breakEvenGainPct, riskRewardRatio: rr,
      };
    }
  }

  // ── Hold zone (-1 to 1) ──────────────────────
  if (pnlPct < -15) {
    return {
      action: 'HOLD — STOP ADDING', urgency: 'MEDIUM', actionColor: '#D68910',
      urgencyBadge: '🟡 MEDIUM',
      timeHorizon: '2–3 months — do not add more capital',
      reasoning: [
        `Neutral technicals (score ${techScore.toFixed(1)}) with ${Math.abs(pnlPct).toFixed(1)}% loss`,
        'Signal is not strong enough to justify averaging — wait for clarity',
        deathCrossBlock ? 'Death Cross active — risk of further downside' : 'Neither bullish nor bearish — momentum unclear',
        `Hard stop loss at ₹${stopLoss} — exit if breached`,
      ],
      stopLoss, targetPrice, averageAt: null, breakEvenGainPct, riskRewardRatio: rr,
    };
  } else if (pnlPct > 20) {
    return {
      action: 'PARTIAL PROFIT', urgency: 'MEDIUM', actionColor: '#D35400',
      urgencyBadge: '🟡 MEDIUM',
      timeHorizon: 'Book 40–50% within 2 weeks, hold rest',
      reasoning: [
        `Up ${pnlPct.toFixed(1)}% with neutral technicals (score ${techScore.toFixed(1)})`,
        'No strong continuation signal — prudent to book partial gains',
        'Hold the rest in case momentum picks up',
        `Protect remaining position with stop at ₹${stopLoss}`,
      ],
      stopLoss, targetPrice, averageAt: null, breakEvenGainPct, riskRewardRatio: rr,
    };
  } else {
    return {
      action: 'HOLD', urgency: 'LOW', actionColor: '#5D6D7E',
      urgencyBadge: '🟢 LOW',
      timeHorizon: '1–3 months — monitor monthly',
      reasoning: [
        `Neutral signal (score ${techScore.toFixed(1)}) — no strong bullish or bearish trigger`,
        pnlPct >= 0 ? `Position profitable at +${pnlPct.toFixed(1)}%` : `Small ${Math.abs(pnlPct).toFixed(1)}% loss — within normal fluctuation`,
        `Target ₹${targetPrice} | Stop Loss ₹${stopLoss}`,
        'Re-evaluate when technical score changes significantly',
      ],
      stopLoss, targetPrice, averageAt: null, breakEvenGainPct, riskRewardRatio: rr,
    };
  }
}

// ─────────────────────────────────────────────
//  Route: Stock Analysis
// ─────────────────────────────────────────────
app.get('/api/analyze/:query', async (req, res) => {
  const raw    = req.params.query;
  console.log('[analyze] Request for:', raw);
  const ticker = resolveTicker(raw);
  console.log('[analyze] Resolved ticker:', ticker);

  let quote = null, history = [];
  try {
    quote = await yf.quote(ticker);
    console.log('[analyze] Quote OK:', quote?.longName);
  } catch (e) {
    console.error('[analyze] Quote error:', e.message);
  }

  try {
    const period1 = new Date();
    period1.setFullYear(period1.getFullYear() - 1);
    console.log('[analyze] Fetching chart...');
    const chart = await yf.chart(ticker, { period1, interval: '1d' });
    history = chart.quotes || [];
    console.log('[analyze] Chart OK, rows:', history.length);
  } catch (e) {
    console.error('[analyze] Chart error:', e.message, e.stack);
    return res.status(404).json({ error: `'${raw}' का data नहीं मिला। Try करें: TCS, RELIANCE, INFY, HDFCBANK` });
  }

  if (history.length < 50) {
    return res.status(404).json({ error: `'${raw}' के लिए पर्याप्त data नहीं। Newly listed stock हो सकता है।` });
  }

  const validHistory = history.filter(d => d.close > 0);
  const closes  = validHistory.map(d => d.close);
  const volumes = validHistory.map(d => d.volume || 0);
  const dates   = validHistory.map(d => new Date(d.date).toISOString().slice(0, 10));
  const opens   = validHistory.map(d => d.open  || d.close);
  const highs   = validHistory.map(d => d.high  || d.close);
  const lows    = validHistory.map(d => d.low   || d.close);

  // Include today's live price so indicators reflect current intraday move
  const livePrice = quote?.regularMarketPrice;
  if (livePrice && livePrice > 0) {
    const todayStr = new Date().toISOString().slice(0, 10);
    if (dates[dates.length - 1] === todayStr) {
      closes[closes.length - 1] = livePrice;
      highs[highs.length - 1]   = Math.max(highs[highs.length - 1], livePrice);
      lows[lows.length - 1]     = Math.min(lows[lows.length - 1],  livePrice);
    } else {
      closes.push(livePrice);
      highs.push(livePrice);
      lows.push(livePrice);
      volumes.push(quote?.regularMarketVolume || 0);
    }
  }

  const ind  = analyse(closes, volumes);
  const name = quote?.longName || quote?.shortName || ticker.replace(/\.(NS|BO)$/, '');
  ind.explanation = buildExplanation(ind, name);

  const cp   = closes[closes.length - 1];
  const prev = closes[closes.length - 2] || cp;

  // Trade Setup — ATR-based SL / Target
  const atr14      = calcATR(highs, lows, closes, 14);
  const score      = ind.totalScore;
  const isBullish  = score >= 1;
  const isBearish  = score <= -1;
  const tradeSetup = {
    atr:          +atr14.toFixed(2),
    entryPrice:   +cp.toFixed(2),
    stopLoss:     isBullish  ? +(cp - 1.5 * atr14).toFixed(2)
                : isBearish  ? +(cp + 1.5 * atr14).toFixed(2) : null,
    target1:      isBullish  ? +(cp + 2 * atr14).toFixed(2)
                : isBearish  ? +(cp - 2 * atr14).toFixed(2) : null,
    target2:      isBullish  ? +(cp + 4 * atr14).toFixed(2)
                : isBearish  ? +(cp - 4 * atr14).toFixed(2) : null,
    riskReward:   '1:2',
    holdingPeriod: score >= 3 ? 'Positional (2–4 weeks)'
                 : score >= 1 ? 'Swing (3–7 days)'
                 : score <= -3 ? 'Exit immediately'
                 : score <= -1 ? 'Exit within 1–2 days'
                 : 'Wait — no clear trade',
    slPct:        isBullish ? +((1.5 * atr14 / cp * 100).toFixed(1)) : null,
    tgt1Pct:      isBullish ? +((2 * atr14 / cp * 100).toFixed(1)) : null,
    tgt2Pct:      isBullish ? +((4 * atr14 / cp * 100).toFixed(1)) : null,
  };

  res.json({
    ticker,
    displayTicker:  ticker.replace(/\.(NS|BO)$/, ''),
    companyName:    name,
    exchange:       ticker.endsWith('.NS') ? 'NSE' : 'BSE',
    sector:         quote?.sector   || 'N/A',
    industry:       quote?.industry || 'N/A',
    currentPrice:   +cp.toFixed(2),
    previousClose:  +prev.toFixed(2),
    priceChange:    +(cp - prev).toFixed(2),
    priceChangePct: +((cp - prev) / prev * 100).toFixed(2),
    marketCap:      quote?.marketCap || null,
    peRatio:        quote?.trailingPE || quote?.forwardPE || null,
    eps:            quote?.epsTrailingTwelveMonths || null,
    week52High:     quote?.fiftyTwoWeekHigh || null,
    week52Low:      quote?.fiftyTwoWeekLow  || null,
    volume:         quote?.regularMarketVolume || null,
    avgVolume:      quote?.averageVolume || null,
    beta:           quote?.beta || null,
    indicators:     ind,
    tradeSetup,
    chartData:      { dates, closes, opens, highs, lows, volumes },
  });
});

// ─────────────────────────────────────────────
//  Route: Bulk Analysis
// ─────────────────────────────────────────────
app.post('/api/bulk', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'File upload नहीं हुई।' });

  let inputs = [];
  try {
    const name = req.file.originalname.toLowerCase();
    if (name.endsWith('.csv')) {
      const lines = req.file.buffer.toString('utf8').split('\n').filter(l => l.trim());
      if (lines.length < 2) return res.status(400).json({ error: 'File खाली है।' });
      const hdr = lines[0].split(',').map(h => h.trim().toLowerCase().replace(/"/g, ''));
      const ti = hdr.findIndex(h => ['ticker','symbol','stock','scrip'].includes(h));
      const ni = hdr.findIndex(h => ['company_name','name','company'].includes(h));
      if (ti < 0 && ni < 0) return res.status(400).json({ error: "File में 'ticker' या 'company_name' column नहीं मिला।" });
      for (let i = 1; i < Math.min(lines.length, 51); i++) {
        const cols = lines[i].split(',').map(c => c.trim().replace(/"/g, ''));
        const v = ti >= 0 ? cols[ti] : cols[ni];
        if (v) inputs.push(v);
      }
    } else {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(req.file.buffer);
      const ws = wb.worksheets[0];
      const hdr = [];
      ws.getRow(1).eachCell(c => hdr.push(String(c.value || '').toLowerCase().trim()));
      const ti = hdr.findIndex(h => ['ticker','symbol','stock','scrip'].includes(h));
      const ni = hdr.findIndex(h => ['company_name','name','company'].includes(h));
      if (ti < 0 && ni < 0) return res.status(400).json({ error: "File में 'ticker' या 'company_name' column नहीं मिला।" });
      ws.eachRow((row, rn) => {
        if (rn === 1 || inputs.length >= 50) return;
        const idx = ti >= 0 ? ti + 1 : ni + 1;
        const v = String(row.getCell(idx).value || '').trim();
        if (v) inputs.push(v);
      });
    }
  } catch (e) {
    return res.status(400).json({ error: 'File parse नहीं हो सकी: ' + e.message });
  }

  const results = [];
  for (const input of inputs) {
    const row = { input, ticker: '', company: input, price: 'N/A', change: 'N/A', marketCap: 'N/A', pe: 'N/A', rsi: 'N/A', macd: 'N/A', maCross: 'N/A', recommendation: 'N/A', confidence: 'N/A', score: 'N/A', error: '' };
    try {
      const ticker = resolveTicker(input);
      row.ticker = ticker.replace(/\.(NS|BO)$/, '');

      let quote = null;
      try { quote = await yf.quote(ticker); } catch (_) {}

      const period1 = new Date(); period1.setFullYear(period1.getFullYear() - 1);
      const chart = await yf.chart(ticker, { period1, interval: '1d' });
      const hist  = chart.quotes || [];
      if (hist.length < 30) throw new Error('Insufficient data');

      const closes  = hist.map(d => d.close || 0).filter(v => v > 0);
      const volumes = hist.map(d => d.volume || 0);
      const liveP = quote?.regularMarketPrice;
      if (liveP && liveP > 0) closes[closes.length - 1] = liveP;
      const ind     = analyse(closes, volumes);

      row.company    = quote?.longName || quote?.shortName || ticker;
      row.price      = '₹' + closes[closes.length - 1].toFixed(2);
      const prev     = closes[closes.length - 2] || closes[closes.length - 1];
      const pct      = ((closes[closes.length - 1] - prev) / prev * 100).toFixed(2);
      row.change     = (pct > 0 ? '+' : '') + pct + '%';
      row.pe         = quote?.trailingPE ? quote.trailingPE.toFixed(1) : 'N/A';
      row.rsi        = ind.rsi.value.toFixed(1);
      row.macd       = ind.macd.score === 1 ? 'Bullish ✅' : 'Bearish ⚠️';
      row.maCross    = ind.movingAverages.crossType === 'golden' ? 'Golden ✅' : ind.movingAverages.crossType === 'death' ? 'Death ⚠️' : 'N/A';
      row.recommendation = ind.recommendation;
      row.confidence = ind.confidence + '%';
      row.score      = (ind.totalScore > 0 ? '+' : '') + ind.totalScore;
      if (quote?.marketCap) {
        const mc = quote.marketCap;
        row.marketCap = mc >= 1e12 ? '₹' + (mc/1e12).toFixed(2) + ' L Cr' : mc >= 1e7 ? '₹' + (mc/1e7).toFixed(2) + ' Cr' : '₹' + (mc/1e5).toFixed(2) + ' L';
      }
    } catch (e) {
      row.error = e.message.includes('Insufficient') ? 'Data कम है' : e.message.slice(0, 60);
    }
    results.push(row);
    await new Promise(r => setTimeout(r, 400));
  }
  res.json({ results });
});

// ─────────────────────────────────────────────
//  Route: Bulk Excel Download
// ─────────────────────────────────────────────
app.post('/api/bulk/download', async (req, res) => {
  const { results } = req.body || {};
  if (!results) return res.status(400).end();
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Stock Analysis');
  const cols = ['Company','Ticker','Price','Change %','Market Cap','P/E','RSI','MACD','MA Cross','Recommendation','Confidence','Score','Error'];
  ws.addRow(cols);
  ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2C3E50' } };
  const cmap = { 'Strong Buy':'FFD5F5E3','Buy':'FFEAFAF1','Hold':'FFFEF9E7','Sell':'FFFDEDEC','Strong Sell':'FFF5B7B1' };
  results.forEach(r => {
    const row = ws.addRow([r.company,r.ticker,r.price,r.change,r.marketCap,r.pe,r.rsi,r.macd,r.maCross,r.recommendation,r.confidence,r.score,r.error]);
    const bg  = cmap[r.recommendation] || 'FFFFFFFF';
    row.eachCell(c => { c.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: bg } }; });
  });
  cols.forEach((_, i) => { ws.getColumn(i + 1).width = 18; });
  const note = ws.addRow(['Generated by Indian Stock Analyzer — Educational use only. Not SEBI-registered advice.']);
  note.font = { italic: true, color: { argb: 'FF999999' }, size: 9 };
  ws.mergeCells(`A${note.number}:M${note.number}`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="stock_analysis.xlsx"');
  await wb.xlsx.write(res);
  res.end();
});

// ─────────────────────────────────────────────
//  Route: Template download
// ─────────────────────────────────────────────
app.get('/api/bulk/template', async (req, res) => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Stocks');
  const h  = ['ticker','company_name','exchange','notes'];
  ws.addRow(h);
  ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF185FA5' } };
  [['TCS','Tata Consultancy Services','NSE','IT'],['HDFCBANK','HDFC Bank','NSE','Banking'],['RELIANCE','Reliance Industries','NSE','Oil & Telecom'],['','Bajaj Finance','NSE','NBFC']].forEach(r => ws.addRow(r));
  h.forEach((_, i) => { ws.getColumn(i + 1).width = 25; });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="template.xlsx"');
  await wb.xlsx.write(res);
  res.end();
});

// ─────────────────────────────────────────────
//  Route: Autocomplete Suggest
// ─────────────────────────────────────────────
app.get('/api/suggest/:q', (req, res) => {
  const q = (req.params.q || '').toLowerCase().trim();
  if (q.length < 1) return res.json([]);
  const results = [];
  for (const [alias, ticker] of Object.entries(ALIASES)) {
    if (alias.includes(q) || ticker.toLowerCase().includes(q)) {
      const display = ticker.replace(/\.(NS|BO)$/, '');
      if (!results.find(r => r.ticker === display)) {
        results.push({ label: alias, ticker: display, fullTicker: ticker });
      }
    }
  }
  res.json(results.slice(0, 8));
});

// ─────────────────────────────────────────────
//  Route: Portfolio Template Download
// ─────────────────────────────────────────────
app.get('/api/portfolio/template', async (req, res) => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('My Portfolio');
  const headers = ['ticker', 'buy_price', 'quantity', 'company_name'];
  ws.addRow(headers);
  ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF185FA5' } };
  const samples = [
    ['TCS',      3500, 10, 'Tata Consultancy Services'],
    ['INFY',     1800, 20, 'Infosys'],
    ['HDFCBANK', 1650, 15, 'HDFC Bank'],
    ['RELIANCE', 2800,  5, 'Reliance Industries'],
    ['WIPRO',     550, 30, 'Wipro'],
  ];
  samples.forEach(r => ws.addRow(r));
  ws.getColumn(1).width = 14;
  ws.getColumn(2).width = 14;
  ws.getColumn(3).width = 12;
  ws.getColumn(4).width = 30;

  // Instructions row
  const instrRow = ws.addRow(['Replace sample data above with your actual holdings. quantity column is optional.']);
  instrRow.font = { italic: true, color: { argb: 'FF888888' }, size: 10 };
  ws.mergeCells(`A${instrRow.number}:D${instrRow.number}`);

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="portfolio_template.xlsx"');
  await wb.xlsx.write(res);
  res.end();
});

// ─────────────────────────────────────────────
//  Route: Portfolio Analysis
// ─────────────────────────────────────────────
app.post('/api/portfolio', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded.' });

  let holdings = [];
  try {
    const name = req.file.originalname.toLowerCase();
    if (name.endsWith('.csv')) {
      const lines = req.file.buffer.toString('utf8').split('\n').filter(l => l.trim());
      if (lines.length < 2) return res.status(400).json({ error: 'File is empty.' });
      const hdr = lines[0].split(',').map(h => h.trim().toLowerCase().replace(/"/g, ''));
      const ti  = hdr.findIndex(h => ['ticker','symbol','stock','scrip'].includes(h));
      const bi  = hdr.findIndex(h => ['buy_price','buy price','price','cost','purchase_price'].includes(h));
      const qi  = hdr.findIndex(h => ['quantity','qty','shares','units'].includes(h));
      const ni  = hdr.findIndex(h => ['company_name','name','company'].includes(h));
      if (ti < 0) return res.status(400).json({ error: "File must have a 'ticker' column." });
      if (bi < 0) return res.status(400).json({ error: "File must have a 'buy_price' column." });
      for (let i = 1; i < Math.min(lines.length, 51); i++) {
        const cols = lines[i].split(',').map(c => c.trim().replace(/"/g, ''));
        const ticker = cols[ti]; const bp = parseFloat(cols[bi]);
        if (!ticker || isNaN(bp) || bp <= 0) continue;
        holdings.push({ ticker, buyPrice: bp, quantity: qi >= 0 ? parseFloat(cols[qi]) || null : null, companyHint: ni >= 0 ? cols[ni] : null });
      }
    } else {
      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(req.file.buffer);
      const ws = wb2.worksheets[0];
      const hdr = [];
      ws.getRow(1).eachCell(c => hdr.push(String(c.value || '').toLowerCase().trim()));
      const ti  = hdr.findIndex(h => ['ticker','symbol','stock','scrip'].includes(h));
      const bi  = hdr.findIndex(h => ['buy_price','buy price','price','cost','purchase_price'].includes(h));
      const qi  = hdr.findIndex(h => ['quantity','qty','shares','units'].includes(h));
      const ni  = hdr.findIndex(h => ['company_name','name','company'].includes(h));
      if (ti < 0) return res.status(400).json({ error: "File must have a 'ticker' column." });
      if (bi < 0) return res.status(400).json({ error: "File must have a 'buy_price' column." });
      ws.eachRow((row, rn) => {
        if (rn === 1 || holdings.length >= 50) return;
        const ticker = String(row.getCell(ti + 1).value || '').trim();
        const bp     = parseFloat(row.getCell(bi + 1).value);
        if (!ticker || isNaN(bp) || bp <= 0) return;
        holdings.push({
          ticker, buyPrice: bp,
          quantity: qi >= 0 ? parseFloat(row.getCell(qi + 1).value) || null : null,
          companyHint: ni >= 0 ? String(row.getCell(ni + 1).value || '').trim() || null : null,
        });
      });
    }
  } catch (e) {
    return res.status(400).json({ error: 'Could not parse file: ' + e.message });
  }

  if (!holdings.length) return res.status(400).json({ error: 'No valid rows found. Check ticker and buy_price columns.' });

  const urgencyOrder = { CRITICAL: 0, HIGH: 1, MEDIUM: 2, OPPORTUNITY: 3, LOW: 4 };
  const results = [];

  for (const h of holdings) {
    const row = {
      input: h.ticker, ticker: '', company: h.companyHint || h.ticker,
      buyPrice: h.buyPrice, currentPrice: null, quantity: h.quantity,
      investedValue: null, currentValue: null, pnlAbs: null, pnlPct: null,
      techScore: null, recommendation: 'N/A', confidence: 'N/A',
      rsi: 'N/A', macd: 'N/A', maCross: 'N/A',
      decision: null, error: '',
    };
    try {
      const ticker = resolveTicker(h.ticker);
      row.ticker = ticker.replace(/\.(NS|BO)$/, '');

      // Fetch all data in parallel for speed
      const period1 = new Date(); period1.setFullYear(period1.getFullYear() - 1);
      const [quote, chart, fsData, searchData] = await Promise.all([
        yf.quote(ticker, {}, { validateResult: false }).catch(() => null),
        yf.chart(ticker, { period1, interval: '1d' }, { validateResult: false }),
        yf.quoteSummary(ticker, { modules: ['financialData'] }, { validateResult: false }).catch(() => null),
        yf.search(ticker, { newsCount: 3 }, { validateResult: false }).catch(() => null),
      ]);

      const hist  = chart.quotes || [];
      if (hist.length < 30) throw new Error('Insufficient historical data');

      const closes  = hist.map(d => d.close || 0).filter(v => v > 0);
      const volumes = hist.map(d => d.volume || 0);
      const liveP2 = quote?.regularMarketPrice;
      if (liveP2 && liveP2 > 0) closes[closes.length - 1] = liveP2;
      const ind     = analyse(closes, volumes);

      const cp = closes[closes.length - 1];
      row.currentPrice  = +cp.toFixed(2);
      row.company       = quote?.longName || quote?.shortName || row.company;
      row.pnlPct        = +((cp - h.buyPrice) / h.buyPrice * 100).toFixed(2);

      if (h.quantity) {
        row.investedValue = +(h.buyPrice  * h.quantity).toFixed(2);
        row.currentValue  = +(cp           * h.quantity).toFixed(2);
        row.pnlAbs        = +(row.currentValue - row.investedValue).toFixed(2);
      }

      row.techScore      = ind.totalScore;
      row.recommendation = ind.recommendation;
      row.confidence     = ind.confidence + '%';
      row.rsi            = ind.rsi.value.toFixed(1);
      row.macd           = ind.macd.score === 1 ? 'Bullish' : 'Bearish';
      row.maCross        = ind.movingAverages.crossType === 'golden' ? 'Golden Cross' : ind.movingAverages.crossType === 'death' ? 'Death Cross' : 'Neutral';

      row.decision = portfolioDecision(ind.totalScore, row.pnlPct, ind, cp, h.buyPrice, quote?.fiftyTwoWeekHigh || null, h.quantity);

      // ── Fundamentals ──────────────────────────
      const fd = fsData?.financialData;
      const w52Low  = quote?.fiftyTwoWeekLow  || null;
      const w52High = quote?.fiftyTwoWeekHigh || null;
      const w52Range = (w52Low && w52High && w52High > w52Low) ? w52High - w52Low : null;
      const w52Pos  = w52Range ? +((cp - w52Low) / w52Range * 100).toFixed(0) : null;
      let w52Label = null;
      if (w52Pos !== null) {
        if      (w52Pos <= 15) w52Label = '🟢 Near 52W Low — Potential Value Zone';
        else if (w52Pos <= 35) w52Label = '🔵 In Lower Range — Watch for entry';
        else if (w52Pos <= 65) w52Label = '⚪ Mid Range';
        else if (w52Pos <= 85) w52Label = '🟡 In Upper Range — Be selective';
        else                   w52Label = '🔴 Near 52W High — Risk of correction';
      }
      row.fundamentals = {
        pe:             quote?.trailingPE         ? +quote.trailingPE.toFixed(1)       : null,
        pb:             quote?.priceToBook        ? +quote.priceToBook.toFixed(2)      : null,
        debtToEquity:   fd?.debtToEquity          ? +fd.debtToEquity.toFixed(1)        : null,
        revenueGrowth:  fd?.revenueGrowth         ? +(fd.revenueGrowth * 100).toFixed(1) : null,
        week52Low:      w52Low  ? +w52Low.toFixed(2)  : null,
        week52High:     w52High ? +w52High.toFixed(2) : null,
        week52Position: w52Pos,
        week52Label:    w52Label,
      };

      // ── News (latest 3 headlines) ─────────────
      row.news = (searchData?.news || []).slice(0, 3).map(n => ({
        title:     n.title     || '',
        publisher: n.publisher || '',
        link:      n.link      || '#',
        age:       n.providerPublishTime
          ? Math.round((Date.now() / 1000 - n.providerPublishTime) / 3600)
          : null,
      }));
    } catch (e) {
      row.error = e.message.includes('Insufficient') ? 'Not enough data' : e.message.slice(0, 80);
      row.decision = { action: 'ERROR', urgency: 'LOW', actionColor: '#888', urgencyBadge: '⚪ N/A', timeHorizon: 'N/A', reasoning: [row.error], stopLoss: null, targetPrice: null, averageAt: null, breakEvenGainPct: 0, riskRewardRatio: 0 };
    }
    results.push(row);
    await new Promise(r => setTimeout(r, 400));
  }

  results.sort((a, b) => (urgencyOrder[a.decision?.urgency] ?? 9) - (urgencyOrder[b.decision?.urgency] ?? 9));

  const validRows = results.filter(r => r.currentPrice !== null);
  const totalInvested    = validRows.filter(r => r.investedValue).reduce((s, r) => s + r.investedValue, 0);
  const totalCurrentVal  = validRows.filter(r => r.currentValue).reduce((s, r) => s + r.currentValue, 0);
  const totalPnlAbs      = totalCurrentVal - totalInvested;
  const totalPnlPct      = totalInvested > 0 ? +(totalPnlAbs / totalInvested * 100).toFixed(2) : null;

  const countByUrgency = {};
  results.forEach(r => { const u = r.decision?.urgency || 'LOW'; countByUrgency[u] = (countByUrgency[u] || 0) + 1; });

  const byPnl = validRows.filter(r => r.pnlPct !== null).sort((a, b) => b.pnlPct - a.pnlPct);
  const summary = {
    totalInvested: totalInvested || null, totalCurrentValue: totalCurrentVal || null,
    totalPnlAbs: totalPnlAbs || null, totalPnlPct,
    countByUrgency,
    bestPerformer:  byPnl[0]  ? { ticker: byPnl[0].ticker,  company: byPnl[0].company,  pnlPct: byPnl[0].pnlPct  } : null,
    worstPerformer: byPnl[byPnl.length-1] ? { ticker: byPnl[byPnl.length-1].ticker, company: byPnl[byPnl.length-1].company, pnlPct: byPnl[byPnl.length-1].pnlPct } : null,
    totalStocks: results.length,
  };

  res.json({ summary, results });
});

// ─────────────────────────────────────────────
//  Route: Portfolio Excel Download
// ─────────────────────────────────────────────
app.post('/api/portfolio/download', async (req, res) => {
  const { results, summary } = req.body || {};
  if (!results) return res.status(400).end();

  const urgencyColors = { CRITICAL: 'FFFFCCCC', HIGH: 'FFFFE0CC', MEDIUM: 'FFFFF3CC', OPPORTUNITY: 'FFCCE5FF', LOW: 'FFCCFFCC' };
  const wb = new ExcelJS.Workbook();

  // ─── Sheet 1: Holdings Summary ───────────────
  const ws1 = wb.addWorksheet('Holdings Summary');
  const cols1 = ['Company','Ticker','Qty','Buy Price','Current Price','Invested (₹)','Current Value (₹)','P&L (₹)','P&L %','Tech Score','RSI','MACD','MA Cross','Action','Urgency','Time Horizon','Stop Loss','Target Price','Avg Price 1','Avg Price 2','Avg Price 3','Risk:Reward','Reasoning'];
  ws1.addRow(cols1);
  ws1.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws1.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2C3E50' } };

  results.forEach(r => {
    const d = r.decision || {};
    const row = ws1.addRow([
      r.company, r.ticker, r.quantity ?? '', r.buyPrice, r.currentPrice ?? '',
      r.investedValue ?? '', r.currentValue ?? '',
      r.pnlAbs ?? '', r.pnlPct !== null ? r.pnlPct + '%' : '',
      r.techScore ?? '', r.rsi, r.macd, r.maCross,
      d.action || '', d.urgency || '', d.timeHorizon || '',
      d.stopLoss ?? '', d.targetPrice ?? '',
      d.averageAt?.[0] ?? '', d.averageAt?.[1] ?? '', d.averageAt?.[2] ?? '',
      d.riskRewardRatio ? d.riskRewardRatio + ':1' : '',
      (d.reasoning || []).join(' | '),
    ]);
    const bg = urgencyColors[d.urgency] || 'FFFFFFFF';
    row.eachCell(c => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }; });
  });
  cols1.forEach((_, i) => { ws1.getColumn(i + 1).width = i === cols1.length - 1 ? 60 : 18; });

  // ─── Sheet 2: Action Plan ─────────────────────
  const ws2 = wb.addWorksheet('Action Plan');
  const cols2 = ['Priority','Company','Ticker','Buy Price','Current Price','P&L%','Action','Time Horizon','Stop Loss','Target','Avg Price 1','Avg Price 2','Avg Price 3','Reasoning'];
  ws2.addRow(cols2);
  ws2.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws2.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2C3E50' } };

  const urgencyRank = { CRITICAL: 1, HIGH: 2, MEDIUM: 3, OPPORTUNITY: 4, LOW: 5 };
  [...results].sort((a, b) => (urgencyRank[a.decision?.urgency] || 9) - (urgencyRank[b.decision?.urgency] || 9)).forEach((r, idx) => {
    const d = r.decision || {};
    const row = ws2.addRow([
      idx + 1, r.company, r.ticker, r.buyPrice, r.currentPrice ?? '',
      r.pnlPct !== null ? r.pnlPct + '%' : '',
      d.action || '', d.timeHorizon || '',
      d.stopLoss ?? '', d.targetPrice ?? '',
      d.averageAt?.[0] ?? '', d.averageAt?.[1] ?? '', d.averageAt?.[2] ?? '',
      (d.reasoning || []).join(' | '),
    ]);
    const bg = urgencyColors[d.urgency] || 'FFFFFFFF';
    row.eachCell(c => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }; });
  });
  cols2.forEach((_, i) => { ws2.getColumn(i + 1).width = i === cols2.length - 1 ? 60 : 18; });

  const note = ws2.addRow([`Generated ${new Date().toLocaleDateString('en-IN')} — Educational only. Not SEBI-registered investment advice.`]);
  note.font = { italic: true, color: { argb: 'FF999999' }, size: 9 };
  ws2.mergeCells(`A${note.number}:N${note.number}`);

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="portfolio_analysis.xlsx"');
  await wb.xlsx.write(res);
  res.end();
});

// ─────────────────────────────────────────────
//  Route: Mutual Fund Search
// ─────────────────────────────────────────────
function fetchJSON(url, opts = {}) {
  return new Promise((resolve, reject) => {
    const parsedUrl = new URL(url);
    const options = {
      hostname: parsedUrl.hostname,
      path: parsedUrl.pathname + parsedUrl.search,
      method: 'GET',
      headers: { 'User-Agent': 'StockAnalyzer/1.0', ...(opts.headers || {}) },
    };
    https.get(options, resp => {
      let body = '';
      resp.on('data', chunk => { body += chunk; });
      resp.on('end', () => { try { resolve(JSON.parse(body)); } catch { resolve(null); } });
    }).on('error', e => reject(e));
  });
}

// ─────────────────────────────────────────────
//  Route: MF Top 20 Picks
// ─────────────────────────────────────────────
// searchQuery = mfapi.in search term to auto-find correct Direct Growth scheme code
// mustContain = keyword that MUST appear in scheme_name to confirm it's the right fund
const TOP_FUNDS = [
  // Large Cap (10)
  { code: 118834, name: 'Mirae Asset Large Cap',        category: 'Large Cap', mustContain: 'mirae' },
  { code: 120503, name: 'Axis Bluechip Fund',            category: 'Large Cap', mustContain: 'axis bluechip' },
  { code: 119761, name: 'HDFC Top 100 Fund',             category: 'Large Cap', mustContain: 'hdfc top 100' },
  { code: 101624, name: 'Canara Robeco Bluechip',        category: 'Large Cap', mustContain: 'canara robeco bluechip' },
  { code: 120586, name: 'Kotak Bluechip Fund',           category: 'Large Cap', mustContain: 'kotak bluechip' },
  { code: 120465, name: 'ICICI Pru Bluechip Fund',       category: 'Large Cap', mustContain: 'icici pru bluechip' },
  { code: 119522, name: 'SBI Bluechip Fund',             category: 'Large Cap', mustContain: 'sbi bluechip' },
  { code: 118701, name: 'Nippon India Large Cap',        category: 'Large Cap', mustContain: 'nippon india large cap' },
  { code: 120684, name: 'UTI Mastershare Fund',          category: 'Large Cap', mustContain: 'uti mastershare' },
  { code: 119230, name: 'DSP Top 100 Equity Fund',       category: 'Large Cap', mustContain: 'dsp top 100' },
  // Mid Cap (6)
  { code: 120828, name: 'Quant Mid Cap Fund',            category: 'Mid Cap',   mustContain: 'quant mid cap' },
  { code: 0,      name: 'HDFC Mid-Cap Opportunities',    category: 'Mid Cap',   mustContain: 'hdfc mid-cap opportunities', searchQuery: 'HDFC Mid-Cap Opportunities Direct Growth' },
  { code: 120594, name: 'Kotak Emerging Equity Fund',    category: 'Mid Cap',   mustContain: 'kotak emerging' },
  { code: 0,      name: 'Nippon India Growth Fund',      category: 'Mid Cap',   mustContain: 'nippon india growth', searchQuery: 'Nippon India Growth Fund Direct Growth' },
  { code: 120505, name: 'Axis Midcap Fund',              category: 'Mid Cap',   mustContain: 'axis midcap' },
  { code: 0,      name: 'SBI Magnum Midcap Fund',        category: 'Mid Cap',   mustContain: 'sbi magnum midcap', searchQuery: 'SBI Magnum Mid Cap Direct Growth' },
  // Small Cap (4)
  { code: 0,      name: 'Nippon India Small Cap',        category: 'Small Cap', mustContain: 'nippon india small cap', searchQuery: 'Nippon India Small Cap Direct Growth' },
  { code: 0,      name: 'SBI Small Cap Fund',            category: 'Small Cap', mustContain: 'sbi small cap', searchQuery: 'SBI Small Cap Fund Direct Growth' },
  { code: 0,      name: 'Axis Small Cap Fund',           category: 'Small Cap', mustContain: 'axis small cap', searchQuery: 'Axis Small Cap Fund Direct Growth' },
  { code: 120829, name: 'Quant Small Cap Fund',          category: 'Small Cap', mustContain: 'quant small cap' },
];

// Resolve correct scheme code: try hardcoded code first, fallback to name search
async function resolveFundCode(fund) {
  // If code provided, verify it matches mustContain
  if (fund.code > 0) {
    const data = await fetchJSON(`https://api.mfapi.in/mf/${fund.code}`);
    if (data?.meta?.scheme_name?.toLowerCase().includes(fund.mustContain)) return { code: fund.code, data };
  }
  // Search by name, find Direct Growth plan
  const query = fund.searchQuery || fund.name;
  const results = await fetchJSON(`https://api.mfapi.in/mf/search?q=${encodeURIComponent(query)}`);
  if (!results?.length) return null;
  // Prefer Direct + Growth plan
  const match = results.find(r => {
    const n = r.schemeName.toLowerCase();
    return n.includes(fund.mustContain) && n.includes('direct') && (n.includes('growth') || n.includes('gr'));
  }) || results.find(r => r.schemeName.toLowerCase().includes(fund.mustContain));
  if (!match) return null;
  const data = await fetchJSON(`https://api.mfapi.in/mf/${match.schemeCode}`);
  return data?.data?.length ? { code: match.schemeCode, data } : null;
}

app.get('/api/mf/top-picks', async (req, res) => {
  const results = await Promise.allSettled(TOP_FUNDS.map(async (fund) => {
    const resolved = await resolveFundCode(fund);
    if (!resolved) return null;
    const { code: resolvedCode, data } = resolved;
    if (!data?.data?.length) return null;

    const navData = [...data.data].reverse(); // oldest → newest
    const navs    = navData.map(d => parseFloat(d.nav)).filter(v => !isNaN(v));
    if (navs.length < 30) return null;
    const currentNav = navs[navs.length - 1];

    const ret = (fromIdx) => {
      const from = navs[fromIdx];
      return from > 0 ? +((currentNav - from) / from * 100).toFixed(2) : null;
    };
    const r1m = ret(Math.max(0, navs.length - 22));
    const r3m = ret(Math.max(0, navs.length - 66));
    const r6m = ret(Math.max(0, navs.length - 130));
    const r1y = ret(Math.max(0, navs.length - 252));
    const r3y = navs.length >= 756  ? ret(Math.max(0, navs.length - 756))  : null;
    const r5y = navs.length >= 1260 ? ret(Math.max(0, navs.length - 1260)) : null;

    // Consistency: positive months over last 12
    let posMonths = 0;
    for (let i = 0; i < 12; i++) {
      const end   = navs[navs.length - 1 - i * 22];
      const start = navs[navs.length - 1 - (i + 1) * 22];
      if (start && end && end > start) posMonths++;
    }

    // Composite score
    const rawScore = (r1y || 0) * 0.40 + (r6m || 0) * 0.25 + (r3m || 0) * 0.20 + (r1m || 0) * 0.15;
    const consistencyBonus = (posMonths / 12) * 5;
    const finalScore = +(rawScore + consistencyBonus).toFixed(1);

    // Confidence % — normalised to 0–100 (25 pts = 100%)
    const confidence = Math.min(100, Math.round(Math.max(0, finalScore) / 25 * 100));

    // 52W stats
    const last252 = navs.slice(-252);
    const w52High = +Math.max(...last252).toFixed(2);
    const w52Low  = +Math.min(...last252).toFixed(2);
    const navFromPeak = +((w52High - currentNav) / w52High * 100).toFixed(1);

    // Signal
    let signal, signalColor;
    if      (finalScore >= 18) { signal = 'Strong Invest';           signalColor = '#00A36C'; }
    else if (finalScore >= 10) { signal = 'Good Entry';              signalColor = '#2ECC71'; }
    else if (finalScore >= 4)  { signal = 'SIP Only — Watch';        signalColor = '#F39C12'; }
    else                       { signal = 'Avoid / Underperforming'; signalColor = '#E74C3C'; }

    // Lumpsum vs SIP
    let entryMode, entryDetail, entryColor;
    if      (navFromPeak < 5)  { entryMode = 'SIP Only';             entryColor = '#185FA5'; entryDetail = 'NAV near 52W high — average in slowly'; }
    else if (navFromPeak < 15) { entryMode = 'SIP + Lumpsum';        entryColor = '#27AE60'; entryDetail = `${navFromPeak}% below peak — partial lumpsum ok`; }
    else                       { entryMode = 'Lumpsum Opportunity';  entryColor = '#8B0000'; entryDetail = `${navFromPeak}% correction — strong value entry`; }

    // NAV history for chart (full history for 3Y/5Y chart support)
    const chartData = navData.map(d => ({ date: d.date, nav: parseFloat(d.nav) }));

    return {
      code: resolvedCode,
      name: data.meta.scheme_name || fund.name,
      shortName: fund.name,
      category: fund.category,
      fundHouse: data.meta.fund_house || '',
      currentNav: +currentNav.toFixed(4),
      r1m, r3m, r6m, r1y, r3y, r5y,
      w52High, w52Low, navFromPeak,
      posMonths, finalScore, confidence,
      signal, signalColor,
      entryMode, entryColor, entryDetail,
      chartData,
    };
  }));

  const picks = results
    .filter(r => r.status === 'fulfilled' && r.value)
    .map(r => r.value)
    .sort((a, b) => b.finalScore - a.finalScore);

  res.json({ picks, timestamp: new Date().toISOString() });
});

app.get('/api/mf/search', async (req, res) => {
  const q = (req.query.q || '').toLowerCase().trim();
  if (q.length < 2) return res.json([]);
  const data = await fetchJSON('https://api.mfapi.in/mf/search?q=' + encodeURIComponent(q));
  if (!data) return res.json([]);
  res.json(data.slice(0, 25));
});

app.get('/api/mf/:code', async (req, res) => {
  const data = await fetchJSON(`https://api.mfapi.in/mf/${req.params.code}`);
  if (!data) return res.status(404).json({ error: 'Fund data नहीं मिला।' });
  const navHistory = (data.data || []).slice(0, 365).map(d => {
    const [dd, mm, yyyy] = d.date.split('-');
    return { date: `${yyyy}-${mm}-${dd}`, nav: parseFloat(d.nav) };
  }).reverse();

  const returns = {};
  if (navHistory.length) {
    const latest = navHistory[navHistory.length - 1].nav;
    const periods = { '1d':1,'1w':7,'1m':30,'3m':91,'6m':182,'1y':365 };
    for (const [key, days] of Object.entries(periods)) {
      const target = new Date(navHistory[navHistory.length - 1].date);
      target.setDate(target.getDate() - days);
      const past = navHistory.filter(n => new Date(n.date) <= target);
      if (past.length) returns[key] = +((latest - past[past.length-1].nav) / past[past.length-1].nav * 100).toFixed(2);
    }
  }
  res.json({ meta: data.meta, navHistory, returns, currentNav: navHistory.length ? navHistory[navHistory.length-1].nav : null });
});

// ─────────────────────────────────────────────
//  Nifty 50 ticker list
// ─────────────────────────────────────────────
const NIFTY50 = [
  'RELIANCE.NS','TCS.NS','HDFCBANK.NS','ICICIBANK.NS','INFY.NS',
  'HINDUNILVR.NS','ITC.NS','SBIN.NS','BHARTIARTL.NS','KOTAKBANK.NS',
  'LT.NS','AXISBANK.NS','ASIANPAINT.NS','MARUTI.NS','HCLTECH.NS',
  'SUNPHARMA.NS','TITAN.NS','BAJFINANCE.NS','WIPRO.NS','ULTRACEMCO.NS',
  'NESTLEIND.NS','POWERGRID.NS','NTPC.NS','TECHM.NS','TATAMOTORS.NS',
  'TATASTEEL.NS','JSWSTEEL.NS','ADANIENT.NS','ADANIPORTS.NS','ONGC.NS',
  'COALINDIA.NS','BAJAJFINSV.NS','DIVISLAB.NS','CIPLA.NS','DRREDDY.NS',
  'EICHERMOT.NS','HEROMOTOCO.NS','BAJAJ-AUTO.NS','BRITANNIA.NS','GRASIM.NS',
  'INDUSINDBK.NS','M&M.NS','TATACONSUM.NS','APOLLOHOSP.NS','BPCL.NS',
  'HINDALCO.NS','VEDL.NS','SBILIFE.NS','HDFCLIFE.NS','SHREECEM.NS'
];

const NIFTY_NEXT50 = [
  'ADANIGREEN.NS','AMBUJACEM.NS','BAJAJHLDNG.NS','BANKBARODA.NS','BERGEPAINT.NS',
  'BOSCHLTD.NS','CHOLAFIN.NS','COLPAL.NS','DMART.NS','GAIL.NS',
  'GODREJCP.NS','GODREJPROP.NS','HAVELLS.NS','HINDPETRO.NS','ICICIPRULI.NS',
  'INDHOTEL.NS','IOC.NS','IRCTC.NS','LICI.NS','LTIM.NS',
  'LUPIN.NS','MUTHOOTFIN.NS','NAUKRI.NS','NMDC.NS','OFSS.NS',
  'PAGEIND.NS','PIDILITIND.NS','PIIND.NS','PFC.NS','RECLTD.NS',
  'SBICARD.NS','SIEMENS.NS','SRF.NS','TORNTPHARM.NS','TRENT.NS',
  'TVSMOTOR.NS','UBL.NS','UPL.NS','VBL.NS','VOLTAS.NS',
  'ZOMATO.NS','PAYTM.NS','NYKAA.NS','DELHIVERY.NS','POLICYBZR.NS',
  'ICICIGI.NS','HDFCAMC.NS','MFSL.NS','ABCAPITAL.NS','SAIL.NS'
];

const NIFTY_MIDCAP100 = [
  'ALKEM.NS','ASTRAL.NS','AUBANK.NS','BALRAMCHIN.NS','BIOCON.NS',
  'CANFINHOME.NS','CESC.NS','CRISIL.NS','DEEPAKNTR.NS','ELGIEQUIP.NS',
  'EXIDEIND.NS','FEDERALBNK.NS','FORTIS.NS','GMRINFRA.NS','GNFC.NS',
  'GRANULES.NS','GUJGASLTD.NS','HAPPSTMNDS.NS','HINDCOPPER.NS','IDFCFIRSTB.NS',
  'INDIGO.NS','IRFC.NS','IEX.NS','JKCEMENT.NS','JSWENERGY.NS',
  'JUBLFOOD.NS','KALYANKJIL.NS','KANSAINER.NS','KEI.NS','KPITTECH.NS',
  'LAURUSLABS.NS','LTTS.NS','MARICO.NS','MCX.NS','METROPOLIS.NS',
  'MGL.NS','MPHASIS.NS','NATCOPHARM.NS','NBCC.NS','NHPC.NS',
  'OBEROIRLTY.NS','PERSISTENT.NS','PETRONET.NS','POLYCAB.NS','RBLBANK.NS',
  'RAMCOCEM.NS','RITES.NS','ROUTE.NS','SCHAEFFLER.NS','SOBHA.NS',
  'SONACOMS.NS','SUNTV.NS','SUPREMEIND.NS','SYNGENE.NS','TATACHEM.NS',
  'TATACOMM.NS','TATAPOWER.NS','THERMAX.NS','TIINDIA.NS','TORNTPOWER.NS',
  'TRIDENT.NS','VGUARD.NS','VINATIORGA.NS','WELCORP.NS','ZEEL.NS',
  'ABFRL.NS','AARTIIND.NS','APLAPOLLO.NS','CAMS.NS','CHOLAHLDNG.NS',
  'CUB.NS','DELTACORP.NS','FINEORG.NS','GLAXO.NS','GODREJIND.NS',
  'GRAPHITE.NS','INTELLECT.NS','ISEC.NS','JUBLPHARMA.NS','LINDEINDIA.NS',
  'LUXIND.NS','MRPL.NS','PFIZER.NS','RADICO.NS','RELAXO.NS',
  'REDINGTON.NS','SEQUENT.NS','STAR.NS','TTKPRESTIG.NS','TEAMLEASE.NS',
  'TIMKEN.NS','UCOBANK.NS','UNIONBANK.NS','NCC.NS','ENGINERSIN.NS',
  'PNBHOUSING.NS','NLCINDIA.NS','APLLTD.NS','NIACL.NS','SHYAMMETL.NS'
];

// ─────────────────────────────────────────────
//  ATR helper
// ─────────────────────────────────────────────
function calcATR(highs, lows, closes, period = 14) {
  const trs = closes.map((c, i) => {
    if (i === 0) return highs[i] - lows[i];
    return Math.max(highs[i] - lows[i], Math.abs(highs[i] - closes[i-1]), Math.abs(lows[i] - closes[i-1]));
  });
  const recent = trs.slice(-period);
  return recent.reduce((a, b) => a + b, 0) / recent.length;
}

// ─────────────────────────────────────────────
//  Black-Scholes Greeks
// ─────────────────────────────────────────────
function erf(x) {
  const t = 1 / (1 + 0.3275911 * Math.abs(x));
  const poly = t * (0.254829592 + t * (-0.284496736 + t * (1.421413741 + t * (-1.453152027 + t * 1.061405429))));
  const result = 1 - poly * Math.exp(-x * x);
  return x >= 0 ? result : -result;
}
function normCDF(x) { return 0.5 * (1 + erf(x / Math.sqrt(2))); }
function normPDF(x) { return Math.exp(-0.5 * x * x) / Math.sqrt(2 * Math.PI); }

function blackScholes(S, K, T, r, sigma, type) {
  if (T <= 0 || sigma <= 0) return { price: 0, delta: 0, gamma: 0, theta: 0, vega: 0 };
  const d1 = (Math.log(S / K) + (r + 0.5 * sigma ** 2) * T) / (sigma * Math.sqrt(T));
  const d2 = d1 - sigma * Math.sqrt(T);
  const nd1 = normPDF(d1);
  if (type === 'call') {
    return {
      price:  +(S * normCDF(d1) - K * Math.exp(-r * T) * normCDF(d2)).toFixed(2),
      delta:  +normCDF(d1).toFixed(4),
      gamma:  +(nd1 / (S * sigma * Math.sqrt(T))).toFixed(6),
      theta:  +(( -(S * nd1 * sigma) / (2 * Math.sqrt(T)) - r * K * Math.exp(-r * T) * normCDF(d2)) / 365).toFixed(4),
      vega:   +(S * nd1 * Math.sqrt(T) / 100).toFixed(4),
    };
  } else {
    return {
      price:  +(K * Math.exp(-r * T) * normCDF(-d2) - S * normCDF(-d1)).toFixed(2),
      delta:  +(normCDF(d1) - 1).toFixed(4),
      gamma:  +(nd1 / (S * sigma * Math.sqrt(T))).toFixed(6),
      theta:  +(( -(S * nd1 * sigma) / (2 * Math.sqrt(T)) + r * K * Math.exp(-r * T) * normCDF(-d2)) / 365).toFixed(4),
      vega:   +(S * nd1 * Math.sqrt(T) / 100).toFixed(4),
    };
  }
}

// ─────────────────────────────────────────────
//  Route: Daily Top Picks (multi-universe scan)
// ─────────────────────────────────────────────
app.get('/api/daily-picks', async (req, res) => {
  const universe = (req.query.universe || 'nifty50').toLowerCase();
  let stockList, maxPicks;
  if (universe === 'nifty200') {
    stockList = [...NIFTY50, ...NIFTY_NEXT50, ...NIFTY_MIDCAP100];
    maxPicks  = 20;
  } else if (universe === 'nifty100') {
    stockList = [...NIFTY50, ...NIFTY_NEXT50];
    maxPicks  = 15;
  } else {
    stockList = NIFTY50;
    maxPicks  = 12;
  }

  const BATCH = 8;
  const results = [];

  for (let i = 0; i < stockList.length; i += BATCH) {
    const batch = stockList.slice(i, i + BATCH);
    const settled = await Promise.allSettled(batch.map(async (ticker) => {
      const period1 = new Date(Date.now() - 400 * 86400000).toISOString().split('T')[0];
      const [quote, hist] = await Promise.all([
        yf.quote(ticker, {}, { validateResult: false }).catch(() => null),
        yf.chart(ticker, { period1, interval: '1d' }, { validateResult: false }).catch(() => null),
      ]);
      if (!quote || !hist) return null;
      const quotes = (hist.quotes || []).filter(q => q.close != null);
      if (quotes.length < 50) return null;

      const closes = quotes.map(q => q.close);
      const highs  = quotes.map(q => q.high || q.close);
      const lows   = quotes.map(q => q.low  || q.close);
      const vols   = quotes.map(q => q.volume || 0);
      const liveP3 = quote?.regularMarketPrice;
      if (liveP3 && liveP3 > 0) closes[closes.length - 1] = liveP3;

      const ind = analyse(closes, vols);
      const atr = calcATR(highs, lows, closes);
      const cp  = closes[closes.length - 1];
      const stopLoss   = +(cp - 1.5 * atr).toFixed(2);
      const target     = +(cp + 3.0 * atr).toFixed(2);
      const riskReward = '1:2';
      const holdingPeriod = ind.totalScore >= 3 ? 'Positional (2–4 weeks)' : 'Swing (3–7 days)';

      const w52High = quote.fiftyTwoWeekHigh || null;
      const w52Low  = quote.fiftyTwoWeekLow  || null;
      const w52Pct  = (w52High && w52Low && w52High > w52Low)
                      ? +((cp - w52Low) / (w52High - w52Low) * 100).toFixed(1) : null;
      const volNow  = quote.regularMarketVolume || 0;
      const volAvg  = quote.averageVolume || 0;
      const volMult = volAvg > 0 ? +(volNow / volAvg).toFixed(1) : null;
      const target2 = +(cp + 4.0 * atr).toFixed(2);

      return {
        ticker: ticker.replace('.NS','').replace('.BO',''),
        fullTicker: ticker,
        company: quote.longName || quote.shortName || ticker,
        sector: quote.sector || null,
        price: +cp.toFixed(2),
        pe: quote.trailingPE ? +quote.trailingPE.toFixed(1) : null,
        week52High: w52High,
        week52Low:  w52Low,
        week52Pct:  w52Pct,
        volumeMult: volMult,
        score: ind.totalScore,
        recommendation: ind.recommendation,
        confidence: ind.confidence,
        atr: +atr.toFixed(2),
        stopLoss,
        target,
        target2,
        riskReward,
        holdingPeriod,
        rsiValue: ind.rsi.value,
        macdSignal: ind.macd.score === 1 ? 'Bullish' : 'Bearish',
        maCross: ind.movingAverages.score === 1 ? 'Golden' : ind.movingAverages.score === -1 ? 'Death' : 'Neutral',
        bbSignal: ind.bollinger.score === 1 ? 'At Support' : ind.bollinger.score === -1 ? 'Overbought' : 'Neutral',
      };
    }));
    settled.forEach(r => { if (r.status === 'fulfilled' && r.value) results.push(r.value); });
  }

  const picks = results
    .filter(r => r.score >= 1.0)
    .sort((a, b) => b.score - a.score)
    .slice(0, maxPicks);

  res.json({ picks, scanned: results.length, universe, timestamp: new Date().toISOString() });
});

// ─────────────────────────────────────────────
//  Route: Company Financials
// ─────────────────────────────────────────────
app.get('/api/financials/:ticker', async (req, res) => {
  const fullTicker = resolveTicker(req.params.ticker);
  try {
    const summary = await yf.quoteSummary(fullTicker, {
      modules: ['incomeStatementHistory', 'balanceSheetHistory', 'cashflowStatementHistory', 'financialData', 'defaultKeyStatistics'],
    }, { validateResult: false });

    const fmtCr = v => v ? +(v / 1e7).toFixed(2) : null; // to Crores
    const fmtPct = v => v ? +(v * 100).toFixed(2) : null;

    const income = (summary.incomeStatementHistory?.incomeStatementHistory || []).map(s => ({
      date: s.endDate ? new Date(s.endDate).getFullYear() : null,
      revenue:     fmtCr(s.totalRevenue),
      grossProfit: fmtCr(s.grossProfit),
      ebitda:      fmtCr(s.ebitda),
      netIncome:   fmtCr(s.netIncome),
      eps:         s.basicEps != null ? +s.basicEps.toFixed(2) : null,
    }));

    const balance = (summary.balanceSheetHistory?.balanceSheetStatements || []).map(s => ({
      date:         s.endDate ? new Date(s.endDate).getFullYear() : null,
      totalAssets:  fmtCr(s.totalAssets),
      totalDebt:    fmtCr(s.totalDebt || s.longTermDebt),
      cash:         fmtCr(s.cash || s.cashAndCashEquivalents),
      bookValue:    fmtCr(s.totalStockholderEquity),
      currentRatio: s.currentRatio != null ? +s.currentRatio.toFixed(2) : null,
    }));

    const cashflow = (summary.cashflowStatementHistory?.cashflowStatements || []).map(s => ({
      date:          s.endDate ? new Date(s.endDate).getFullYear() : null,
      operatingCF:   fmtCr(s.totalCashFromOperatingActivities),
      investingCF:   fmtCr(s.totalCashflowsFromInvestingActivities),
      financingCF:   fmtCr(s.totalCashFromFinancingActivities),
      freeCF:        s.totalCashFromOperatingActivities && s.capitalExpenditures
                       ? fmtCr(s.totalCashFromOperatingActivities + s.capitalExpenditures)
                       : null,
    }));

    const fd = summary.financialData || {};
    const ks = summary.defaultKeyStatistics || {};
    const ratios = {
      peRatio:       fd.forwardPE ? +fd.forwardPE.toFixed(2) : null,
      pbRatio:       ks.priceToBook ? +ks.priceToBook.toFixed(2) : null,
      roe:           fmtPct(fd.returnOnEquity),
      roa:           fmtPct(fd.returnOnAssets),
      debtToEquity:  fd.debtToEquity ? +fd.debtToEquity.toFixed(2) : null,
      revenueGrowth: fmtPct(fd.revenueGrowth),
      earningsGrowth:fmtPct(fd.earningsGrowth),
      operatingMargin: fmtPct(fd.operatingMargins),
      profitMargin:  fmtPct(fd.profitMargins),
      currentRatio:  fd.currentRatio ? +fd.currentRatio.toFixed(2) : null,
      quickRatio:    fd.quickRatio ? +fd.quickRatio.toFixed(2) : null,
      freeCashflow:  fmtCr(fd.freeCashflow),
    };

    res.json({ ticker: req.params.ticker, income, balance, cashflow, ratios });
  } catch (e) {
    res.status(500).json({ error: 'Financial data not available for this stock.' });
  }
});

// ─────────────────────────────────────────────
//  Helpers for Index Option Signal
// ─────────────────────────────────────────────
const NSE_INDEX_SYMBOLS = { '^NSEI': 'NIFTY', '^NSEBANK': 'BANKNIFTY' };

// Calculate upcoming expiry dates without NSE API
// NIFTY: every Tuesday (NSE changed from Thu) | BANKNIFTY: every Wednesday
function getNearestExpiries(symbol, count = 6) {
  const expiryDay = symbol === 'BANKNIFTY' ? 3 : 2; // 3=Wed, 2=Tue
  const expiries = [];
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  let checked = 0;
  // Check today first — if today is expiry day, include it
  if (d.getDay() === expiryDay) {
    expiries.push(d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }).replace(/ /g, '-'));
  }
  while (expiries.length < count && checked < 90) {
    d.setDate(d.getDate() + 1);
    checked++;
    if (d.getDay() === expiryDay) {
      expiries.push(d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }).replace(/ /g, '-'));
    }
  }
  return expiries;
}

function parseExpiryDate(label) {
  // '07-Apr-2025' → Date
  return new Date(label.replace(/-/g, ' '));
}

// Generate NSE trading symbol for Zerodha Kite
// Weekly:  NIFTY2540723000CE  (YY + single-digit-month + DD + strike + type)
// Monthly: NIFTY25APR23000CE  (YY + MMM + strike + type)
function generateNseSymbol(indexName, expiryLabel, strike, type) {
  const d = parseExpiryDate(expiryLabel);
  const yy = String(d.getFullYear()).slice(2);
  const month = d.getMonth(); // 0-indexed
  const dd = String(d.getDate()).padStart(2, '0');
  const monthAbbr = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'][month];
  // Weekly month chars: 1-9 = Jan-Sep, O = Oct, N = Nov, D = Dec
  const weeklyM = ['1','2','3','4','5','6','7','8','9','O','N','D'][month];

  // Determine if this is the monthly expiry (last Tuesday/Wednesday of month)
  const nextWeek = new Date(d);
  nextWeek.setDate(d.getDate() + 7);
  const isMonthly = nextWeek.getMonth() !== d.getMonth();

  const sym = isMonthly
    ? `${indexName}${yy}${monthAbbr}${strike}${type}`   // e.g. NIFTY25APR23000CE
    : `${indexName}${yy}${weeklyM}${dd}${strike}${type}`; // e.g. NIFTY2540723000CE
  return sym;
}

// ─────────────────────────────────────────────
//  Route: Options Chain + Greeks
// ─────────────────────────────────────────────
app.get('/api/options/:ticker', async (req, res) => {
  const fullTicker = resolveTicker(req.params.ticker);
  const expiry = req.query.expiry || null; // e.g. '07-Apr-2025'
  const R = 0.065;

  // ── NSE index options (NIFTY / BANKNIFTY) ──
  // NSE blocks server-side API calls. We compute everything from Yahoo Finance:
  // spot + India VIX + 60d OHLCV → technical analysis → Black-Scholes signal.
  const nseSymbol = NSE_INDEX_SYMBOLS[fullTicker];
  if (nseSymbol) {
    try {
      const sixtyDaysAgo = new Date(Date.now() - 60 * 86400000).toISOString().split('T')[0];
      const [quoteData, vixData, chartData] = await Promise.all([
        yf.quote(fullTicker, {}, { validateResult: false }).catch(() => null),
        yf.quote('^INDIAVIX', {}, { validateResult: false }).catch(() => null),
        yf.chart(fullTicker, { period1: sixtyDaysAgo, interval: '1d' }, { validateResult: false }).catch(() => null),
      ]);

      const spot = quoteData?.regularMarketPrice || 0;
      const change = quoteData?.regularMarketChange || 0;
      const changePct = quoteData?.regularMarketChangePercent || 0;
      const vix = vixData?.regularMarketPrice || 15;
      const sigma = vix / 100; // e.g. VIX 14 → IV 14%

      // Technical analysis from historical data
      const quotes = chartData?.quotes || [];
      const closes  = quotes.map(q => q.close).filter(Boolean);
      const volumes = quotes.map(q => q.volume || 0);
      const ind = closes.length >= 14 ? analyse(closes, volumes) : null;
      const score = ind?.totalScore || 0;

      // Direction
      let direction, dirEmoji, dirColor;
      if      (score >= 2)  { direction = 'Bullish';  dirEmoji = '📈'; dirColor = '#3B6D11'; }
      else if (score <= -2) { direction = 'Bearish';  dirEmoji = '📉'; dirColor = '#A32D2D'; }
      else                  { direction = 'Neutral';  dirEmoji = '↔️'; dirColor = '#854F0B'; }

      // Strike rounding
      const step = nseSymbol === 'BANKNIFTY' ? 100 : 50;
      const atmStrike = Math.round(spot / step) * step;

      // Expiries
      const expiries = getNearestExpiries(nseSymbol);
      const selectedExpiry = expiry || expiries[0];
      const expiryDate = parseExpiryDate(selectedExpiry);
      const T = Math.max((expiryDate - Date.now()) / (365 * 86400000), 0.001);

      // Recommended option
      let recType, recStrike;
      if (direction === 'Bullish') {
        recType = 'CE'; recStrike = atmStrike;
      } else if (direction === 'Bearish') {
        recType = 'PE'; recStrike = atmStrike;
      } else {
        recType = 'CE+PE'; recStrike = atmStrike; // strangle
      }

      // NSE trading symbols for Kite
      const nseSymbolCE = generateNseSymbol(nseSymbol, expiries[0], recStrike, 'CE');
      const nseSymbolPE = generateNseSymbol(nseSymbol, expiries[0], recStrike, 'PE');
      const recNseSymbol = recType === 'PE' ? nseSymbolPE : nseSymbolCE;

      // Chain: ATM ± 5 strikes
      const strikes = [];
      for (let i = -5; i <= 5; i++) strikes.push(atmStrike + i * step);

      const chain = strikes.map(K => {
        const call = blackScholes(spot, K, T, R, sigma, 'call');
        const put  = blackScholes(spot, K, T, R, sigma, 'put');
        return {
          strike: K,
          callPrice:  call.price  != null ? +call.price.toFixed(2)  : null,
          putPrice:   put.price   != null ? +put.price.toFixed(2)   : null,
          callDelta:  call.delta  != null ? +call.delta.toFixed(3)  : null,
          putDelta:   put.delta   != null ? +put.delta.toFixed(3)   : null,
          gamma:      call.gamma  != null ? +call.gamma.toFixed(6)  : null,
          callTheta:  call.theta  != null ? +call.theta.toFixed(2)  : null,
          putTheta:   put.theta   != null ? +put.theta.toFixed(2)   : null,
          vega:       call.vega   != null ? +call.vega.toFixed(2)   : null,
          iv:         +(sigma * 100).toFixed(1),
          isATM: K === atmStrike,
        };
      });

      // ATR for index-level SL/Target
      const highs  = quotes.map(q => q.high).filter(Boolean);
      const lows   = quotes.map(q => q.low).filter(Boolean);
      const atr    = highs.length >= 14 ? calcATR(highs, lows, closes, 14) : spot * 0.01;

      // Lot sizes (as per NSE)
      const lotSize = nseSymbol === 'BANKNIFTY' ? 15 : 25;

      // Entry / SL / Targets
      const atmRow  = chain.find(c => c.isATM);
      const rawEntry = recType === 'PE' ? atmRow?.putPrice : atmRow?.callPrice;
      const entryLow  = rawEntry ? +rawEntry.toFixed(0) : 0;
      const entryHigh = rawEntry ? +(rawEntry * 1.08).toFixed(0) : 0;

      // Premium-level SL/Target
      const premiumSL   = rawEntry ? +(rawEntry * 0.40).toFixed(0) : 0;  // exit at -60%
      const premiumTgt1 = rawEntry ? +(rawEntry * 1.80).toFixed(0) : 0;  // +80%
      const premiumTgt2 = rawEntry ? +(rawEntry * 2.50).toFixed(0) : 0;  // +150%

      // Index-level SL/Target (ATR-based)
      const isBullish = direction === 'Bullish';
      const isBearish = direction === 'Bearish';
      const indexSL      = isBullish ? +(spot - 1.5 * atr).toFixed(0)
                         : isBearish ? +(spot + 1.5 * atr).toFixed(0) : null;
      const indexTarget1 = isBullish ? +(spot + 1.5 * atr).toFixed(0)
                         : isBearish ? +(spot - 1.5 * atr).toFixed(0) : null;
      const indexTarget2 = isBullish ? +(spot + 3.0 * atr).toFixed(0)
                         : isBearish ? +(spot - 3.0 * atr).toFixed(0) : null;

      // Breakeven index level
      const breakeven = recType === 'CE' ? recStrike + entryLow
                      : recType === 'PE' ? recStrike - entryLow : null;

      // Lot economics
      const costPerLot   = entryLow * lotSize;
      const maxLossLot   = (entryLow - premiumSL) * lotSize;
      const profitLot1   = (premiumTgt1 - entryLow) * lotSize;
      const profitLot2   = (premiumTgt2 - entryLow) * lotSize;
      const thetaPerDay  = atmRow ? Math.abs(recType === 'PE' ? atmRow.putTheta : atmRow.callTheta) : 0;
      const thetaCostLot = +(thetaPerDay * lotSize).toFixed(0);

      const holdingPeriod = score >= 3 ? 'Hold till expiry' : score >= 1 || score <= -1 ? '1–2 days' : 'Intraday only';

      // Signal pills
      const signals = [];
      if (ind) {
        signals.push(ind.rsi.score >= 1 ? `✅ RSI ${ind.rsi.value} — Oversold (Bullish)` : ind.rsi.score <= -1 ? `⚠️ RSI ${ind.rsi.value} — Overbought (Bearish)` : `➖ RSI ${ind.rsi.value} — Neutral`);
        signals.push(ind.macd.score === 1 ? '✅ MACD Bullish crossover' : '⚠️ MACD Bearish signal');
        if (ind.movingAverages.score !== 0) signals.push(ind.movingAverages.crossType === 'golden' ? '✅ Golden Cross — Uptrend' : '⚠️ Death Cross — Downtrend');
        if (ind.bollinger.score === 1) signals.push('✅ Price near lower Bollinger — Bounce possible');
        if (ind.bollinger.score === -1) signals.push('⚠️ Price near upper Bollinger — Overbought');
        if (ind.volume.score === 0.5) signals.push('✅ Volume spike — Momentum confirmed');
      }

      return res.json({
        nseIndex:       true,
        symbol:         nseSymbol,
        spot:           +spot.toFixed(2),
        change:         +change.toFixed(2),
        changePct:      +changePct.toFixed(2),
        vix:            +vix.toFixed(2),
        direction, dirEmoji, dirColor,
        score:          +score.toFixed(2),
        confidence:     ind?.confidence || 0,
        recommendation: ind?.recommendation || 'Hold',
        recType, recStrike,
        recNseSymbol, nseSymbolCE, nseSymbolPE,
        entryLow, entryHigh,
        premiumSL, premiumTgt1, premiumTgt2,
        indexSL, indexTarget1, indexTarget2,
        breakeven,
        lotSize, costPerLot, maxLossLot, profitLot1, profitLot2,
        thetaPerDay: +thetaPerDay.toFixed(2), thetaCostLot,
        atr: +atr.toFixed(0),
        holdingPeriod,
        expiries,
        selectedExpiry,
        chain,
        signals,
        links: [
          { label: '🔗 NSE Option Chain', url: 'https://www.nseindia.com/option-chain', note: 'Select ' + nseSymbol },
          { label: '🔗 Sensibull', url: 'https://sensibull.com/nifty', note: 'Real-time Greeks' },
          { label: '🔗 Opstra', url: 'https://opstra.definedge.com/', note: 'OI Analysis' },
        ],
      });
    } catch (e) {
      return res.status(500).json({ error: 'Index signal failed: ' + e.message });
    }
  }

  // ── Stock options via Yahoo Finance ──
  try {
    const [quoteData, optData] = await Promise.all([
      yf.quote(fullTicker, {}, { validateResult: false }).catch(() => null),
      expiry
        ? yf.options(fullTicker, { date: new Date(expiry) }, { validateResult: false }).catch(() => null)
        : yf.options(fullTicker, {}, { validateResult: false }).catch(() => null),
    ]);

    if (!optData) return res.status(404).json({ error: 'Options data not available for this stock. Note: NSE stock F&O data may be limited on Yahoo Finance.' });

    const spot = quoteData?.regularMarketPrice || quoteData?.ask || 0;
    const expiries = optData.expirationDates || [];
    const selectedExpiry = expiry || (expiries[0] ? new Date(expiries[0] * 1000).toISOString().split('T')[0] : null);
    const T = selectedExpiry ? Math.max((new Date(selectedExpiry) - Date.now()) / (365 * 86400000), 0.001) : 0.05;

    const processContracts = (contracts = [], type) =>
      contracts.map(c => {
        const iv = c.impliedVolatility || 0.3;
        const K  = c.strike;
        const greeks = spot > 0 && K > 0 ? blackScholes(spot, K, T, R, iv, type) : {};
        const isATM = Math.abs(K - spot) <= spot * 0.02;
        return {
          strike: K, type, lastPrice: c.lastPrice, bid: c.bid, ask: c.ask,
          volume: c.volume, openInterest: c.openInterest, iv: +(iv * 100).toFixed(1),
          delta: greeks.delta, gamma: greeks.gamma, theta: greeks.theta, vega: greeks.vega,
          isATM, inTheMoney: c.inTheMoney,
          expiry: c.expiration ? new Date(c.expiration * 1000).toISOString().split('T')[0] : selectedExpiry,
        };
      });

    const calls = processContracts(optData.options?.[0]?.calls, 'call');
    const puts  = processContracts(optData.options?.[0]?.puts,  'put');

    const atmCall = calls.find(c => c.isATM) || calls[Math.floor(calls.length / 2)];
    const avgIV = atmCall?.iv || 30;
    let strategy = '', strategyReason = '';
    if (!quoteData) {
      strategy = 'N/A'; strategyReason = 'Could not determine underlying trend.';
    } else {
      const pct52 = quoteData.fiftyTwoWeekHigh && quoteData.fiftyTwoWeekLow
        ? (spot - quoteData.fiftyTwoWeekLow) / (quoteData.fiftyTwoWeekHigh - quoteData.fiftyTwoWeekLow) * 100 : 50;
      const bullish = pct52 > 55, bearish = pct52 < 40, highIV = avgIV > 35;
      if (bullish && !highIV)      { strategy = '📈 Buy Call';         strategyReason = `Bullish (${pct52.toFixed(0)}% of 52w) + Low IV (${avgIV}%) → directional call.`; }
      else if (bullish && highIV)  { strategy = '📊 Bull Call Spread'; strategyReason = `Bullish + High IV (${avgIV}%) → spread reduces cost.`; }
      else if (bearish && !highIV) { strategy = '📉 Buy Put';          strategyReason = `Bearish (${pct52.toFixed(0)}% of 52w) + Low IV → directional put.`; }
      else if (bearish && highIV)  { strategy = '🐻 Bear Put Spread';  strategyReason = `Bearish + High IV → spread reduces cost.`; }
      else if (highIV)             { strategy = '🎯 Short Strangle';   strategyReason = `Sideways + High IV (${avgIV}%) → sell OTM call + put.`; }
      else                         { strategy = '⏸️ Wait & Watch';     strategyReason = `No strong signal. IV moderate (${avgIV}%).`; }
    }

    res.json({
      ticker: req.params.ticker, spot,
      expiries: expiries.map(e => new Date(e * 1000).toISOString().split('T')[0]),
      selectedExpiry, calls: calls.slice(0, 40), puts: puts.slice(0, 40),
      strategy, strategyReason, avgIV, source: 'Yahoo Finance',
    });
  } catch (e) {
    res.status(500).json({ error: 'Options data not available: ' + e.message });
  }
});

// ─────────────────────────────────────────────
//  Global error handler
// ─────────────────────────────────────────────
app.use((err, req, res, next) => {
  console.error('[ERROR HANDLER]', err.message, err.stack);
  res.status(500).json({ error: 'Server error: ' + err.message });
});

// ─────────────────────────────────────────────
//  Serve frontend for all other routes
// ─────────────────────────────────────────────
app.use((req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const PORT = process.env.PORT || 3000;
// Local server (ignored on Vercel)
if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`\n✅ Indian Stock Analyzer running!`);
    console.log(`🌐 Open: http://localhost:${PORT}\n`);
  });
}

// Required for Vercel serverless
module.exports = app;
