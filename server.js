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
  'nifty':'%5ENSEI','sensex':'%5EBSESN',
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

  let rec, color, emoji;
  if      (total >= 3)  { rec = 'Strong Buy';  color = '#00A36C'; emoji = '🚀'; }
  else if (total >= 1)  { rec = 'Buy';          color = '#2ECC71'; emoji = '✅'; }
  else if (total >= -1) { rec = 'Hold';         color = '#F39C12'; emoji = '⏸️'; }
  else if (total >= -3) { rec = 'Sell';         color = '#E74C3C'; emoji = '⚠️'; }
  else                  { rec = 'Strong Sell';  color = '#8B0000'; emoji = '🔴'; }

  const confidence = Math.min(100, Math.round(Math.abs(total) / 5.5 * 100));

  return {
    rsi:            { value: +rsiVal.toFixed(2), score: rsiScore },
    macd:           { macd: +macdVal.toFixed(4), signal: +signalVal.toFixed(4), histogram: +histVal.toFixed(4), score: macdScore },
    movingAverages: { sma50: sma50 ? +sma50.toFixed(2) : null, sma200: sma200 ? +sma200.toFixed(2) : null, score: maScore, crossType: (sma50 && sma200) ? (sma50 > sma200 ? 'golden' : 'death') : 'neutral' },
    bollinger:      { upper: +bbU.toFixed(2), middle: +bbM.toFixed(2), lower: +bbL.toFixed(2), score: bbScore },
    volume:         { avg5d: Math.round(vol5), avg20d: Math.round(vol20), score: volScore },
    totalScore:     +total.toFixed(2),
    recommendation: rec, color, emoji, confidence,
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

  let t = `${name}: ${ind.recommendation} — ${ind.confidence}% confidence.`;
  if (pos.length) t += '\n✅ Bullish signals: ' + pos.join(', ') + '.';
  if (neg.length) t += '\n⚠️ Bearish signals: ' + neg.join(', ') + '.';
  t += '\n\n⚡ Disclaimer: Educational only. Consult a SEBI-registered advisor before investing.';
  return t;
}

// ─────────────────────────────────────────────
//  Portfolio Decision Engine
// ─────────────────────────────────────────────
function portfolioDecision(techScore, pnlPct, ind, currentPrice, buyPrice, week52High) {
  const sma200 = ind.movingAverages.sma200;
  const crossType = ind.movingAverages.crossType;
  const rsi = ind.rsi.value;

  // ── Computed values ──────────────────────────
  const stopLoss = +Math.max(buyPrice * 0.88, currentPrice * 0.93).toFixed(2);

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
      return {
        action: 'CUT LOSS', urgency: 'HIGH', actionColor: '#C0392B',
        urgencyBadge: '🟠 HIGH',
        timeHorizon: 'Exit within 1 week',
        reasoning: [
          `Bearish technical signal (score ${techScore.toFixed(1)}) combined with ${Math.abs(pnlPct).toFixed(1)}% loss`,
          'Risk of further decline is high — technicals confirm weakness',
          deathCrossBlock ? 'Death Cross active — long-term trend is negative' : `Stop loss at ₹${stopLoss} already breached`,
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

  const closes  = history.map(d => d.close  || 0).filter(v => v > 0);
  const volumes = history.map(d => d.volume || 0);
  const dates   = history.map(d => new Date(d.date).toISOString().slice(0, 10));
  const opens   = history.map(d => d.open  || d.close || 0);
  const highs   = history.map(d => d.high  || d.close || 0);
  const lows    = history.map(d => d.low   || d.close || 0);

  const ind  = analyse(closes, volumes);
  const name = quote?.longName || quote?.shortName || ticker.replace(/\.(NS|BO)$/, '');
  ind.explanation = buildExplanation(ind, name);

  const cp   = closes[closes.length - 1];
  const prev = closes[closes.length - 2] || cp;

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

      let quote = null;
      try { quote = await yf.quote(ticker); } catch (_) {}

      const period1 = new Date(); period1.setFullYear(period1.getFullYear() - 1);
      const chart = await yf.chart(ticker, { period1, interval: '1d' });
      const hist  = chart.quotes || [];
      if (hist.length < 30) throw new Error('Insufficient historical data');

      const closes  = hist.map(d => d.close || 0).filter(v => v > 0);
      const volumes = hist.map(d => d.volume || 0);
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

      row.decision = portfolioDecision(ind.totalScore, row.pnlPct, ind, cp, h.buyPrice, quote?.fiftyTwoWeekHigh || null);
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
function fetchJSON(url) {
  return new Promise((resolve) => {
    https.get(url, { headers: { 'User-Agent': 'StockAnalyzer/1.0' } }, resp => {
      let body = '';
      resp.on('data', chunk => { body += chunk; });
      resp.on('end', () => { try { resolve(JSON.parse(body)); } catch { resolve(null); } });
    }).on('error', () => resolve(null));
  });
}

app.get('/api/mf/search', async (req, res) => {
  const q = (req.query.q || '').toLowerCase().trim();
  if (q.length < 2) return res.json([]);
  const data = await fetchJSON('https://api.mfapi.in/mf');
  if (!data) return res.json([]);
  const starts   = data.filter(f => f.schemeName.toLowerCase().startsWith(q));
  const contains = data.filter(f => f.schemeName.toLowerCase().includes(q) && !f.schemeName.toLowerCase().startsWith(q));
  res.json([...starts, ...contains].slice(0, 25));
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
app.listen(PORT, () => {
  console.log(`\n✅ Indian Stock Analyzer चालू है!`);
  console.log(`🌐 Browser में खोलें: http://localhost:${PORT}\n`);
});
