// ============================================================================
//  levels/level6.js — Level 6 "THE WITCH'S RECKONING": the final boss + ending.
//
//  Summoned through the Door of Light, the white-haired King stands on a lava
//  shore. The WITCH rises giant from the lava on the left and hovers over an
//  unreachable molten river. She cannot be struck directly while a secondary
//  boss stands with the King on the cavern floor — she summons the six-armed
//  GUARDIAN (weaker than in the Black Halls), and the only way to wound her then
//  is to BLOCK a scimitar and send it flying back into her. Fell the guardian
//  and she is briefly open to the King's own lava bullets — until she conjures
//  the LAVA KNIGHT, and the cycle begins again, round after round, until her
//  power is spent. Her death quakes the cavern; the magic carpet appears and
//  bears the King RIGHT-TO-LEFT across the lava river to the hole he once fell
//  through, out through the quaking castle, and up into the clouds of the peak
//  where — tired and sick — he lies back and lets the carpet carry him home.
//
//  Reuses the L5 lava-cave look (drawBackground5), the L3 guardian art
//  (drawBoss / drawScimitar / drawFlyingSword), the L5 Lava Knight art
//  (drawKnight5 + position-relative hitboxes), the biters as flying heads, the
//  flying carpet, the L1 sky for the ending, and the shared text primitives.
//  See plans/modularization-refactor.md.
// ============================================================================
'use strict';

// ---- arena geometry (bounds live here; the plats/lava live in 03-level-data.js)
const ARENA_L6 = 18700, ARENA_R6 = 20300;    // where the King + a summoned boss may stand
const WITCH_X6 = 18250, WITCH_Y6 = FLOOR6 - 250;   // she hovers over the shore lava, on the left
const LANE6 = { low: FLOOR6 - 12, mid: FLOOR6 - 44, high: FLOOR6 - 84 };
const SWORD_REACH6 = 900;                     // how far the guardian's boomerangs fly before returning
const WITCH_HP = 6;                           // total power (≈ 3 cycles)

const l6 = {
  phase: 'summon',        // summon → fight → death → carpet → flight → descend → castle → clouds → credits
  t: 0, sub: 0,
  witch: null,
  guardian: null, knight: null,
  round: 0, bossType: null,   // 'guardian' | 'knight'
  vulnT: 0,                   // >0 while the witch is open to the King's lava bullets
  bullets: [],                // the King's lava bullets (fired with the Fire-Sword)
  balls: [],                  // cosmetic lava splashes for the arena
  quakeT: 0, quakeMag: 0,
  flash: 0,
  flight: null,
  msg: '', msgT: 0,
  dialog: { text: '', t: 0, dur: 0 }, dialogDelay: null,
  end: { t: 0 },
  credits: 0,
};

function l6toast(s) { l6.msg = s; l6.msgT = 3.4; }
function l6say(text, dur) { l6.dialog = { text: text, t: 0, dur: dur || 5 }; }

// a decaying camera quake (the engine's draw applies l6.quake* as a cam offset)
function l6quake(mag, dur) { l6.quakeMag = Math.max(l6.quakeMag, mag); l6.quakeT = Math.max(l6.quakeT, dur); }

function lavaAt6(x) {
  for (const L of LAVA6) if (x > L.x0 && x < L.x1) return L;
  return null;
}

function initEntsL6() {
  l6.phase = 'summon'; l6.t = 0; l6.sub = 0;
  l6.witch = { x: WITCH_X6, y: WITCH_Y6, hp: WITCH_HP, maxHp: WITCH_HP,
    appearT: 0, hurtT: 0, dead: false, deadT: 0, cackleT: 0 };
  l6.guardian = null; l6.knight = null;
  l6.round = 0; l6.bossType = null; l6.vulnT = 0;
  l6.bullets = []; l6.balls = [];
  l6.quakeT = 0; l6.quakeMag = 0; l6.flash = 0;
  l6.flight = null;
  l6.msg = ''; l6.msgT = 0;
  l6.dialog = { text: '', t: 0, dur: 0 }; l6.dialogDelay = null;
  l6.end = { t: 0 }; l6.credits = 0;
}

// ============================================================================
//  THE GIANT WITCH — an enlarged, animated take on the L3 hooded witch, given
//  hp and a state (rising, hovering, hurt-flash, dying). Drawn in world space.
// ============================================================================
function drawGiantWitch() {
  const w = l6.witch;
  if (!w) return;
  const rise = smooth(clamp(w.appearT, 0, 1));
  if (rise <= 0) return;
  const dead = w.dead ? clamp(1 - w.deadT * 0.5, 0, 1) : 1;
  if (dead <= 0) return;
  const S = 3.4;                                   // she is GIANT
  const bob = Math.sin(T * 1.1) * 10;
  const hurt = w.hurtT > 0 ? (Math.floor(T * 30) % 2 === 0) : false;
  const vuln = l6.vulnT > 0;                        // glows red-open when reachable
  lg.push();
  lg.translate(w.x, w.y + (1 - rise) * 260 + bob);  // rises up out of the lava
  lg.scale(S, S * rise);
  // cold aura (or a hot vulnerable aura)
  if (vuln) { lg.setColor(1.0, 0.5, 0.3, 0.16 + 0.06 * Math.sin(T * 8)); }
  else { lg.setColor(0.5, 0.85, 0.9, 0.10 * dead); }
  lg.circle('fill', 0, -10, 78);
  // flowing robe
  const rc = hurt ? [0.7, 0.2, 0.2] : [0.06, 0.05, 0.10];
  lg.setColor(rc[0], rc[1], rc[2], dead);
  const sway = Math.sin(T * 1.4) * 6;
  lg.polygon('fill', -26, 54, 26, 54, 14 + sway, -34, -14 + sway, -34);
  lg.setColor(rc[0] * 1.5, rc[1] * 1.5, rc[2] * 1.6, dead);   // inner robe fold
  lg.polygon('fill', -12, 50, 12, 50, 7 + sway, -30, -7 + sway, -30);
  // long raised arms conjuring
  const ra = Math.sin(T * 2.0) * 5;
  lg.setColor(rc[0], rc[1], rc[2], dead); lg.setLineWidth(6);
  lg.line(-10, -20, -34, -50 - ra); lg.line(10, -20, 34, -50 + ra);
  // clawed hands
  lg.setColor(0.5, 0.9, 0.85, dead * 0.8);
  lg.circle('fill', -34, -50 - ra, 3); lg.circle('fill', 34, -50 + ra, 3);
  // hood + head
  lg.setColor(0.04, 0.03, 0.06, dead);
  lg.circle('fill', 0, -42, 13);
  lg.polygon('fill', -15, -34, 15, -34, 0, -66);
  // glowing eyes
  const eg = 0.6 + 0.4 * Math.sin(T * 5);
  lg.setColor(vuln ? 1.0 : 0.6, vuln ? 0.5 : 1.0, vuln ? 0.4 : 0.9, dead * eg);
  lg.circle('fill', -4, -44, 2.0); lg.circle('fill', 4, -44, 2.0);
  // crooked staff with a glowing orb
  lg.setColor(0.3, 0.22, 0.14, dead); lg.setLineWidth(3.4);
  lg.line(34, -50, 40, 42);
  const og = 0.7 + 0.3 * Math.sin(T * 7);
  lg.setColor(vuln ? 1.0 : 0.6, vuln ? 0.55 : 1.0, vuln ? 0.35 : 0.9, dead * og);
  lg.circle('fill', 34, -56, 7);
  lg.setColor(1, 1, 1, dead * og * 0.6); lg.circle('fill', 33, -57, 3);
  lg.setLineWidth(1);
  lg.pop();
}

// world-space hitbox around the witch's body (only meaningful while reachable)
function witchHitbox() {
  const w = l6.witch;
  return { x1: w.x - 70, y1: w.y - 230, x2: w.x + 70, y2: w.y + 190 };
}
function pointInWitch(x, y) {
  const b = witchHitbox();
  return x > b.x1 && x < b.x2 && y > b.y1 && y < b.y2;
}

function hurtWitch(n) {
  const w = l6.witch;
  if (!w || w.dead) return;
  w.hp -= n; w.hurtT = 0.35;
  l6.flash = Math.max(l6.flash, 0.2);
  if (sfxHit) sfxHit.play(0.6, 0.55 + love.math.random() * 0.1);
  spawnL6Splash(w.x, w.y - 40, 8);
  if (w.hp <= 0) { w.hp = 0; beginWitchDeath(); }
  else { l6toast('The witch shrieks — her power wanes  (' + w.hp + ')'); }
}

// ============================================================================
//  SECONDARY BOSSES — the witch conjures one at a time onto the cavern floor.
//  Rendering reuses the L3 guardian art and the L5 knight art via a temporary
//  alias of l3.boss / l5.knight (both are inactive during Level 6).
// ============================================================================
function spawnGuardian6() {
  l6.bossType = 'guardian';
  l6.guardian = {
    x: ARENA_L6 + 220, y: FLOOR6, hp: 6, active: true, hitCool: 0, appearT: 0,
    swords: [], fireCool: 1.6, order: [], dead: false, deadT: 0,
    armSwing: 0, touchCool: 0, arms: [true, true, true, true, true, true], throwArm: -1,
  };
  l6say('The witch summons her six-armed guardian — turn its own blade against her!', 5.5);
}

function spawnKnight6() {
  l6.bossType = 'knight';
  l6.knight = {
    x: ARENA_R6 - 120, y: FLOOR6, dir: -1, hp: 4, state: 'gallop',
    active: true, dead: false, deadT: 0, hitCool: 0, ph: 0, flash: 0, swing: 0,
    volley: 3, fireCool: 2.2, pauseT: 0, bolts: [],
  };
  l6say('The Lava Knight rides from the fire — cut it down before you strike her again!', 5.5);
}

function drawL6Guardian() {
  if (!l6.guardian) return;
  const save = l3.boss; l3.boss = l6.guardian; drawBoss(); l3.boss = save;
}
function drawL6Knight() {
  if (!l6.knight) return;
  const save = l5.knight; l5.knight = l6.knight; drawKnight5(); l5.knight = save;
}

// the shuffled low/mid/high volley, local to the L6 guardian
function nextLane6(b) {
  if (!b.order.length) {
    b.order = ['low', 'mid', 'high'];
    for (let i = b.order.length - 1; i > 0; i--) {
      const j = Math.floor(love.math.random() * (i + 1));
      const t = b.order[i]; b.order[i] = b.order[j]; b.order[j] = t;
    }
  }
  return b.order.pop();
}
function pickArm6(b, dir) {
  const pref = dir > 0 ? [3, 4, 5] : [2, 1, 0];
  for (const i of pref) if (b.arms[i]) return i;
  for (let i = 0; i < 6; i++) if (b.arms[i]) return i;
  return -1;
}

// register a melee hit on the guardian (called from the swing window)
function tryHitGuardian6(p) {
  const b = l6.guardian;
  if (!b || b.dead || !b.active || b.hitCool > 0) return false;
  if (Math.abs(p.x - b.x) > 64 || p.facing !== (b.x < p.x ? -1 : 1)) return false;
  if (Math.abs(p.y - b.y) > 70) return false;
  b.hp -= 1; b.hitCool = 0.6;
  if (sfxHit) sfxHit.play(0.6, 0.7 + love.math.random() * 0.12);
  const away = (p.x >= b.x) ? 1 : -1;
  p.vx = away * 420; p.vy = -180; p.state = 'air'; p.t = 0; p.inv = Math.max(p.inv || 0, 0.3);
  spawnDust(b.x + away * 30, b.y - 40, 10, 1.3);
  if (b.hp <= 0) { defeatSecondaryBoss('guardian'); }
  else { l6toast('Guardian struck!  ' + b.hp + ' blow' + (b.hp === 1 ? '' : 's') + ' remain'); }
  return true;
}

function updateGuardian6(dt, p) {
  const b = l6.guardian;
  if (!b) return;
  b.appearT = Math.min(1, b.appearT + dt * 1.2);
  b.hitCool = Math.max(0, b.hitCool - dt);
  b.armSwing = Math.max(0, b.armSwing - dt);
  b.touchCool = Math.max(0, b.touchCool - dt);
  if (b.dead) { b.deadT += dt; return; }
  // stalk toward the King, kept on the arena floor
  const want = clamp(p.x, ARENA_L6 + 120, ARENA_R6 - 160);
  const step = 30 * dt;
  if (want > b.x + 2) b.x = Math.min(want, b.x + step);
  else if (want < b.x - 2) b.x = Math.max(want, b.x - step);
  // body contact hurts the King
  if (!p.dying && (p.inv || 0) <= 0 && b.touchCool <= 0 && Math.abs(p.x - b.x) < 32 && p.y > FLOOR6 - 150) {
    b.touchCool = 0.6; hurtPlayer(p, p.x >= b.x ? 1 : -1); spawnDust(p.x, p.y - 30, 6, 0.9);
  }
  // fire boomerang scimitars, one lane at a time, up to three in flight
  b.fireCool -= dt;
  if (b.fireCool <= 0 && b.swords.length < 3) {
    const lane = nextLane6(b);
    const dir = (p.x >= b.x) ? 1 : -1;
    const armIndex = pickArm6(b, dir);
    if (armIndex >= 0) b.arms[armIndex] = false;
    b.swords.push({ x: b.x + dir * 40, y: LANE6[lane], vx: 560 * dir, dir: dir, lane: lane, phase: 'out', spin: 0, armIndex: armIndex, deflected: false });
    b.fireCool = b.order.length ? 0.95 : 1.8;
    b.armSwing = 0.35; b.throwArm = armIndex;
    if (sfxSwing) sfxSwing.play(0.4, 0.7 + love.math.random() * 0.1);
  }
  // move swords; a DEFLECTED blade flies on toward the witch instead of returning
  for (let i = b.swords.length - 1; i >= 0; i--) {
    const s = b.swords[i];
    s.spin += dt * 15 * (s.dir || 1);
    s.x += s.vx * dt;
    if (s.deflected) {
      // hurled back past the guardian, streaking toward the witch on the left
      if (s.x <= l6.witch.x + 30) {
        hurtWitch(1);
        spawnL6Splash(s.x, s.y, 6);
        b.swords.splice(i, 1); continue;
      }
      if (s.x < ARENA_L6 - 1600) { b.swords.splice(i, 1); continue; }
    } else if (s.phase === 'out') {
      if (Math.abs(s.x - b.x) >= SWORD_REACH6) { s.phase = 'back'; s.vx = -560 * s.dir; }
    } else {
      if ((s.dir || 1) > 0 ? s.x <= b.x + 20 : s.x >= b.x - 20) {
        if (s.armIndex >= 0) b.arms[s.armIndex] = true;
        b.swords.splice(i, 1); continue;
      }
    }
    // contact with the King — the MID lane may be blocked (→ flung at the witch)
    if (!p.dying && Math.abs(s.x - p.x) < 22 && !s.deflected) {
      const top = heroTop(p), bot = p.y;
      if (s.y + 9 > top && s.y - 9 < bot) {
        const dir = s.vx > 0 ? 1 : -1;
        if (s.lane === 'mid' && (p.blockT || 0) > 0 && p.facing === -dir) {
          s.deflected = true; s.vx = -Math.abs(560) * (dir > 0 ? 1 : 1);   // send it toward the witch (left)
          s.vx = -560; s.dir = -1;
          if (s.armIndex >= 0) b.arms[s.armIndex] = true;   // his hand is free again
          p.blockFlash = 0.25;
          if (sfxParry) sfxParry.play(0.55, 1.0 + love.math.random() * 0.12);
          spawnDust(p.x + dir * 10, p.y - 30, 6, 0.8);
          l6toast('Blocked!  The blade streaks back at the witch');
        } else if ((p.inv || 0) <= 0) {
          hurtPlayer(p, dir); spawnDust(p.x, p.y - 30, 5, 0.8);
        }
      }
    }
  }
}

// The L6 knight: same art + position-relative hitboxes as L5, but patrolling the
// arena bounds and with no fire-sword drop on death.
function tryHitKnight6(p) {
  const k = l6.knight;
  if (!k || k.dead || !k.active || k.hitCool > 0) return false;
  const swordBox = heroSwordHitBox(p);
  let touched = false;
  for (const box of lavaKnightHitBoxes(k)) { if (rectsOverlap(swordBox, box)) { touched = true; break; } }
  if (!touched) return false;
  k.hp -= 1; k.hitCool = 0.5; k.flash = 0.3;
  if (sfxHit) sfxHit.play(0.6, 0.85 + love.math.random() * 0.1);
  const away = (p.x >= k.x) ? 1 : -1;
  p.vx = away * 300; p.vy = -150; p.state = 'air'; p.t = 0; p.inv = Math.max(p.inv || 0, 0.35);
  spawnDust(k.x, k.y - 60, 8, 1.1);
  if (k.hp <= 0) { defeatSecondaryBoss('knight'); }
  else { l6toast('Lava Knight struck!  ' + k.hp + ' blow' + (k.hp === 1 ? '' : 's') + ' remain'); }
  return true;
}

function updateKnight6(dt, p) {
  const k = l6.knight;
  if (!k) return;
  k.flash = Math.max(0, k.flash - dt);
  k.hitCool = Math.max(0, k.hitCool - dt);
  if (k.dead) { k.deadT += dt; return; }
  const speed = 128 + (4 - k.hp) * 12;
  k.x += k.dir * speed * dt;
  k.ph += speed * dt * 0.02;
  if (k.x < ARENA_L6 + 40) { k.x = ARENA_L6 + 40; k.dir = 1; }
  else if (k.x > ARENA_R6 - 40) { k.x = ARENA_R6 - 40; k.dir = -1; }
  if (!p.dying && (p.inv || 0) <= 0 && k.hitCool <= 0 && Math.abs(p.x - k.x) < 48 && p.y > FLOOR6 - 74) {
    k.hitCool = 0.7; k.swing = 0.3; hurtPlayer(p, k.dir); spawnDust(p.x, p.y - 30, 6, 1.0);
  }
  k.swing = Math.max(0, k.swing - dt);
  if (k.pauseT > 0) {
    k.pauseT -= dt;
    if (k.pauseT <= 0) { k.volley = 3; k.fireCool = 0.3; }
  } else {
    k.fireCool -= dt;
    if (k.fireCool <= 0 && k.volley > 0) {
      const ox = k.x + k.dir * 18, oy = FLOOR6 - 96;
      const tx = p.x, ty = p.y - 30;
      const dx = tx - ox, dy = ty - oy, d = Math.hypot(dx, dy) || 1;
      k.bolts.push({ x: ox, y: oy, vx: dx / d * 400, vy: dy / d * 400, t: 0, r: 7 });
      k.swing = 0.3; if (sfxSwing) sfxSwing.play(0.4, 0.7);
      k.volley -= 1; k.fireCool = 0.55;
      if (k.volley <= 0) k.pauseT = 3.6;
    }
  }
  for (let i = k.bolts.length - 1; i >= 0; i--) {
    const b = k.bolts[i];
    b.t += dt; b.x += b.vx * dt; b.y += b.vy * dt;
    if (b.t > 2.6 || b.y > FLOOR6 + 40 || b.x < ARENA_L6 - 500 || b.x > ARENA_R6 + 500) { k.bolts.splice(i, 1); continue; }
    if (!p.dying && Math.abs(b.x - p.x) < b.r + 11 && b.y > heroTop(p) && b.y < p.y) {
      const dir = b.vx > 0 ? 1 : -1;
      if ((p.blockT || 0) > 0 && p.facing === -dir) {
        p.blockFlash = 0.25; if (sfxParry) sfxParry.play(0.45, 0.85);
        spawnL6Splash(b.x, b.y, 3); k.bolts.splice(i, 1); continue;
      }
      if ((p.inv || 0) <= 0) { hurtPlayer(p, dir); spawnL6Splash(b.x, b.y, 4); k.bolts.splice(i, 1); }
    }
  }
}

// a secondary boss falls → the witch is briefly open to the King's lava bullets
function defeatSecondaryBoss(kind) {
  const b = kind === 'guardian' ? l6.guardian : l6.knight;
  if (b) { b.dead = true; b.deadT = 0; b.active = false; if (b.swords) b.swords.length = 0; if (b.bolts) b.bolts.length = 0; }
  if (kind === 'guardian') for (let i = 0; i < 6; i++) l6.guardian.arms[i] = true;
  l6.vulnT = 6.0;   // window to fire lava bullets at the witch
  l6.bossType = null;
  spawnDust(b ? b.x : WITCH_X6, FLOOR6 - 40, 16, 1.4);
  if (l6.witch.hp > 0) l6toast('The witch is exposed!  Loose your fire at her — quickly!');
}

// ============================================================================
//  LAVA SPLASHES (arena-local, so they draw during Level 6)
// ============================================================================
function spawnL6Splash(x, y, n) {
  if (l6.balls.length > 140) return;
  for (let i = 0; i < n; i++) {
    l6.balls.push({ x: x + (love.math.random() - 0.5) * 24, y: y - 4,
      vx: (love.math.random() - 0.5) * 260, vy: -(120 + love.math.random() * 240),
      r: 3 + love.math.random() * 4, t: 0 });
  }
}
function updateL6Balls(dt) {
  for (let i = l6.balls.length - 1; i >= 0; i--) {
    const b = l6.balls[i];
    b.t += dt; b.vy += 1200 * dt; b.x += b.vx * dt; b.y += b.vy * dt;
    if (b.y > FLOOR6 + 40 || b.t > 2) l6.balls.splice(i, 1);
  }
}

// ============================================================================
//  WITCH DEATH → QUAKE → CARPET
// ============================================================================
function beginWitchDeath() {
  const w = l6.witch;
  w.dead = true; w.deadT = 0;
  l6.phase = 'death'; l6.t = 0;
  l6.flash = 0.7; l6quake(26, 3.2);
  spawnDust(w.x, w.y - 30, 26, 2.2); spawnL6Splash(w.x, w.y, 24);
  if (sfxThunder) sfxThunder.play(0.7, 0.6);
  l6toast('The witch is undone!');
}

// ============================================================================
//  CARPET FLIGHT — right-to-left across the lava river to the hole he fell from.
//  Mirrors the L5 flight (which runs left-to-right): the auto-scroll is negative,
//  the camera leads to the LEFT, heads sweep in from ahead (the left), and lava
//  bolts rise from the river. It ends at the hole on the far-left lip.
// ============================================================================
const FL6 = { TOP: 96, BOT: 400, RIVER: 452, ALT: 236, CAMY: 252, SCROLL: 262, VFLY: 340, HFLY: 175 };

function startFlight6() {
  const p = player;
  l6.flight = {
    active: true, phase: 'lift', t: 0,
    heads: [], upBolts: [], headCool: 1.2, boltCool: 1.0,
    startX: p.x, y0: p.y, holeX: HOLE_X6, blackA: 0,
  };
  p.state = 'cine'; p.vx = 0; p.vy = 0; p.facing = -1;
  p.sheathed = false; p.swordIdle = 0;
  p.lavaSword = true; p.lavaCharge = 3;
  p.hp = difficultyMaxHp(); p.inv = 0.6; p.blockHold = 0;
  l6.bullets.length = 0;
  l6.phase = 'flight';
}

function flight6Hurt(p) {
  const f = l6.flight;
  if (!f) return;
  if (IMMORTAL) { p.inv = Math.max(p.inv || 0, 0.4); return; }
  if ((p.inv || 0) > 0 || p.dying) return;
  p.hp = (p.hp || difficultyMaxHp()) - 1; p.inv = 1.1; p.blockFlash = 0.2;
  spawnDust(p.x, p.y, 5, 0.9); if (sfxHit) sfxHit.play(0.5, 1.0);
  if (p.hp <= 0) { p.hp = difficultyMaxHp(); p.inv = 1.6; p.lavaCharge = 3; l6toast('Hold on — stay aloft!'); }
}

function updateFlight6Ents(dt) {
  const f = l6.flight, p = player;
  const au = 1 - (p.atkT || 0) / ATK_DUR;
  const swordActive = (p.atkT || 0) > 0 && au > 0.30 && au < 0.62;
  for (let i = f.heads.length - 1; i >= 0; i--) {
    const h = f.heads[i];
    h.t += dt;
    if (h.state === 'dead') { h.dead += dt; if (h.dead > 0.5) f.heads.splice(i, 1); continue; }
    h.x += h.vx * dt;
    h.y = clamp(h.y + h.vy * dt + Math.sin((T + h.ph) * 3) * 26 * dt, FL6.TOP, FL6.BOT);
    if (h.x > cam.x + VW * 0.72) { f.heads.splice(i, 1); continue; }
    if (swordActive) {
      const dx = h.x - p.x;
      if (dx * p.facing > 0 && Math.abs(dx) < 70 && Math.abs(h.y - (p.y - 28)) < 56) {
        h.state = 'dead'; h.dead = 0; spawnDust(h.x, h.y, 7, 1.0); continue;
      }
    }
    if ((p.inv || 0) <= 0 && Math.abs(h.x - p.x) < 24 && Math.abs(h.y - (p.y - 18)) < 24) {
      flight6Hurt(p); h.state = 'dead'; h.dead = 0;
    }
  }
  for (let i = f.upBolts.length - 1; i >= 0; i--) {
    const b = f.upBolts[i];
    b.t += dt; b.vy += 55 * dt; b.x += b.vx * dt; b.y += b.vy * dt;
    if (b.y < FL6.TOP - 90 || b.t > 4.5) { f.upBolts.splice(i, 1); continue; }
    if ((p.inv || 0) <= 0 && Math.abs(b.x - p.x) < b.r + 12 && Math.abs(b.y - (p.y - 18)) < b.r + 18) {
      flight6Hurt(p); f.upBolts.splice(i, 1);
    }
  }
  for (let i = l6.bullets.length - 1; i >= 0; i--) {
    const bu = l6.bullets[i];
    bu.t += dt; bu.x += bu.vx * dt; bu.y += bu.vy * dt;
    let gone = bu.t > 1.6 || bu.x < cam.x - VW * 0.62;
    for (const h of f.heads) {
      if (h.state === 'dead') continue;
      if (Math.abs(bu.x - h.x) < 24 && Math.abs(bu.y - h.y) < 24) {
        h.state = 'dead'; h.dead = 0; gone = true; spawnDust(h.x, h.y, 6, 0.9);
        if (sfxHit) sfxHit.play(0.5, 1.2);
      }
    }
    if (gone) l6.bullets.splice(i, 1);
  }
}

function updateFlight6(dt) {
  const f = l6.flight, p = player;
  f.t += dt;
  p.inv = Math.max(0, (p.inv || 0) - dt);
  p.blockFlash = Math.max(0, (p.blockFlash || 0) - dt);
  p.blockT = Math.max(0, (p.blockT || 0) - dt);
  p.atkT = Math.max(-1, (p.atkT || 0) - dt);
  p.drawT = Math.max(0, (p.drawT || 0) - dt);
  p.lavaCharge = p.lavaCharge || 0;
  p.state = 'ground'; p.onGround = true; p.vx = 0; p.facing = -1;

  if (f.phase === 'lift') {
    const k = smooth(clamp(f.t / 1.3, 0, 1));
    p.y = lerp(f.y0, FL6.ALT, k);
    p.x = f.startX - f.t * 140;
    cam.x = lerp(cam.x, p.x - 190, Math.min(1, dt * 3)); cam.y = lerp(cam.y, FL6.CAMY, Math.min(1, dt * 3)); cam.zoom = 1;
    if (f.t > 1.3) { f.phase = 'run'; f.t = 0; }
    return;
  }
  if (f.phase === 'run') {
    updateFireCharge(p, dt);
    const up = keyUp(), down = keyDown(), left = keyLeft(), right = keyRight();
    let vy = 0; if (up) vy -= FL6.VFLY; if (down) vy += FL6.VFLY;
    p.y = clamp(p.y + vy * dt, FL6.TOP, FL6.BOT);
    let vx = -FL6.SCROLL; if (left) vx -= FL6.HFLY; if (right) vx += FL6.HFLY * 0.8;
    p.x += vx * dt;
    cam.x = lerp(cam.x, p.x - 190, Math.min(1, dt * 4)); cam.y = FL6.CAMY; cam.zoom = 1;
    // flying heads sweep in from AHEAD (the left)
    f.headCool -= dt;
    if (f.headCool <= 0) {
      f.headCool = 0.7 + love.math.random() * 0.85;
      const hy = FL6.TOP + 24 + love.math.random() * (FL6.BOT - FL6.TOP - 48);
      f.heads.push({ x: cam.x - VW * 0.60, y: hy, vx: (135 + love.math.random() * 80),
        vy: (love.math.random() - 0.5) * 46, ph: love.math.random() * 6, t: 0,
        phase: love.math.random() * 6.28, state: 'chase', bite: 0, hurt: 0, dead: 0 });
    }
    // lava bolts rising from the river
    f.boltCool -= dt;
    if (f.boltCool <= 0) {
      f.boltCool = 0.5 + love.math.random() * 0.65;
      const bx = p.x + (love.math.random() - 0.3) * 340;
      f.upBolts.push({ x: bx, y: FL6.RIVER, vx: (love.math.random() - 0.5) * 40,
        vy: -(255 + love.math.random() * 130), r: 7, t: 0 });
    }
    updateFlight6Ents(dt);
    if (p.x <= f.holeX + 120) { f.phase = 'descend'; f.t = 0; }
    return;
  }
  if (f.phase === 'descend') {
    // the carpet dives into the hole he once fell through; fade to black
    p.x = lerp(p.x, f.holeX, Math.min(1, dt * 2));
    p.y += 180 * dt;
    cam.x = lerp(cam.x, f.holeX, Math.min(1, dt * 3)); cam.y = lerp(cam.y, p.y - 80, Math.min(1, dt * 3));
    f.blackA = Math.min(1, (f.blackA || 0) + dt * 0.8);
    if (f.t > 2.2) { l6.phase = 'castle'; l6.t = 0; f.active = false; l6.flight.blackA = 1; }
    return;
  }
}

// ============================================================================
//  DRAW — arena, river, witch, bosses, projectiles, and the flight entities
// ============================================================================
function drawLava6() {
  const camL = cam.x - VW * 0.62 / cam.zoom - 40, camR = cam.x + VW * 0.62 / cam.zoom + 40;
  for (const L of LAVA6) {
    if (L.x1 < camL || L.x0 > camR) continue;
    const x0 = Math.max(L.x0, camL), x1 = Math.min(L.x1, camR), w = x1 - x0;
    if (w <= 0) continue;
    const DEPTH = 1600, N = 26;
    for (let i = 0; i < N; i++) {
      const k = i / N;
      lg.setColor(lerp(0.55, 0.055, k), lerp(0.14, 0.02, k), lerp(0.05, 0.02, k), 1);
      lg.rectangle('fill', x0, L.y + i * (DEPTH / N), w, DEPTH / N + 1);
    }
    lg.setColor(0.9, 0.32, 0.07, 1); lg.rectangle('fill', x0, L.y, w, 40);
    lg.setColor(1.0, 0.62, 0.14, 0.95);
    const step = 22;
    for (let x = Math.floor(x0 / step) * step; x < x1; x += step) {
      if (x < L.x0) continue;
      const yy = L.y + Math.sin(x * 0.05 + T * 3) * 4 + Math.sin(x * 0.13 + T * 5) * 2;
      lg.rectangle('fill', Math.max(x, x0), yy, Math.min(step, x1 - Math.max(x, x0)), 6);
    }
    lg.setColor(1.0, 0.5, 0.12, 0.10); lg.rectangle('fill', x0, L.y - 50, w, 50);
  }
}

function drawL6Balls() {
  for (const b of l6.balls) {
    lg.setColor(1.0, 0.5, 0.12, 0.18); lg.circle('fill', b.x, b.y, b.r * 2.1);
    lg.setColor(0.95, 0.32, 0.06, 1); lg.circle('fill', b.x, b.y, b.r);
    lg.setColor(1.0, 0.82, 0.3, 1); lg.circle('fill', b.x - b.r * 0.25, b.y - b.r * 0.25, b.r * 0.5);
  }
}

function drawL6Bullets() {
  for (const bu of l6.bullets) {
    lg.setColor(1.0, 0.45, 0.1, 0.18); lg.circle('fill', bu.x - bu.vx * 0.012, bu.y - bu.vy * 0.012, bu.r * 1.6);
    lg.setColor(1.0, 0.55, 0.12, 0.4); lg.circle('fill', bu.x, bu.y, bu.r * 1.7);
    lg.setColor(1.0, 0.35, 0.08, 1); lg.circle('fill', bu.x, bu.y, bu.r);
    lg.setColor(1.0, 0.92, 0.5, 1); lg.circle('fill', bu.x - 1.5, bu.y - 1.5, bu.r * 0.45);
  }
}

function drawEntsL6() {
  drawLava6();
  drawGiantWitch();
  if (l6.bossType === 'guardian' || (l6.guardian && (!l6.guardian.dead || l6.guardian.deadT < 1))) drawL6Guardian();
  if (l6.bossType === 'knight' || (l6.knight && (!l6.knight.dead || l6.knight.deadT < 1.4))) drawL6Knight();
  // the knight's flung lava bolts
  if (l6.knight && l6.knight.bolts) {
    for (const b of l6.knight.bolts) {
      lg.setColor(1.0, 0.45, 0.1, 0.22); lg.circle('fill', b.x - b.vx * 0.012, b.y - b.vy * 0.012, b.r * 1.7);
      lg.setColor(0.95, 0.32, 0.06, 1); lg.circle('fill', b.x, b.y, b.r);
      lg.setColor(1.0, 0.82, 0.35, 1); lg.circle('fill', b.x - b.r * 0.3, b.y - b.r * 0.3, b.r * 0.5);
    }
  }
  drawL6Balls();
  drawL6Bullets();
  if (l6.flight && l6.flight.active) drawFlight6Ents();
}

function drawFlight6Ents() {
  const f = l6.flight;
  if (!f || !f.active) return;
  for (const b of f.upBolts) {
    lg.setColor(1.0, 0.45, 0.1, 0.22); lg.circle('fill', b.x - b.vx * 0.01, b.y - b.vy * 0.01, b.r * 1.8);
    lg.setColor(0.95, 0.32, 0.06, 1); lg.circle('fill', b.x, b.y, b.r);
    lg.setColor(1.0, 0.82, 0.35, 1); lg.circle('fill', b.x - b.r * 0.3, b.y - b.r * 0.3, b.r * 0.5);
  }
  for (const h of f.heads) drawBiter(h);
}

// dim red cave backdrop (reuses the L5 look, plus the far witch-shore glow)
function drawBackground6(cam) { drawBackground5(cam); }

// ============================================================================
//  THE ENDING — castle escape → clouds → credits (cinematic beats)
// ============================================================================
function updateCastle6(dt) {
  l6.t += dt;
  l6quake(14, 0.4);   // the castle keeps shuddering from the quake
  // the carpet drifts the King leftward across a shaking hall, then out
  const p = player;
  p.state = 'cine'; p.facing = -1;
  if (l6.t < 1.0) { /* arrive out of black */ }
  p.x = lerp(p.x, p.x - 260 * dt, 1);
  p.y = FLOOR6 - 120 + Math.sin(T * 1.4) * 8;
  cam.x = p.x; cam.y = FLOOR6 - 180; cam.zoom = 1;
  if (l6.t > 5.5) { l6.phase = 'clouds'; l6.t = 0; l6.end.t = 0; }
}

function updateClouds6(dt) {
  l6.t += dt;
  const p = player;
  p.state = 'cine'; p.facing = 1;
  // he lies back on the carpet and lets it carry him; the camera holds on sky
  p.x = 0; p.y = 260 + Math.sin(T * 0.8) * 10;
  cam.x = 0; cam.y = 180; cam.zoom = 1;
  l6.end.t += dt;
  // credit cards appear one after another
  if (l6.end.t > 3.0 && l6.credits < 1) l6.credits = 1;
  if (l6.end.t > 6.5 && l6.credits < 2) l6.credits = 2;
  if (l6.end.t > 10.0 && l6.credits < 3) l6.credits = 3;
}

let CLOUDS6 = null;
function drawClouds6() {
  // a soft dusk sky with slow-drifting cloud banks
  for (let i = 0; i <= 16; i++) {
    const k = i / 16;
    lg.setColor(lerp(0.29, 0.62, k), lerp(0.31, 0.52, k), lerp(0.5, 0.55, k), 1);
    lg.rectangle('fill', 0, VH * k, VW, VH / 16 + 1);
  }
  // sun glow low on the horizon
  lg.setColor(1.0, 0.86, 0.6, 0.18); lg.circle('fill', VW * 0.5, VH * 0.62, 220);
  lg.setColor(1.0, 0.9, 0.7, 0.5); lg.circle('fill', VW * 0.5, VH * 0.62, 70);
  if (!CLOUDS6) {
    CLOUDS6 = [];
    const rng = love.math.newRandomGenerator(97);
    for (let i = 0; i < 10; i++) {
      CLOUDS6.push({ x: rng.random() * VW, y: VH * (0.18 + rng.random() * 0.5),
        w: 180 + rng.random() * 300, h: 22 + rng.random() * 30, spd: 5 + rng.random() * 10,
        a: 0.5 + rng.random() * 0.35 });
    }
  }
  for (const c of CLOUDS6) {
    const x = ((c.x - T * c.spd) % (VW + c.w * 2) + (VW + c.w * 2)) % (VW + c.w * 2) - c.w;
    lg.setColor(0.96, 0.92, 0.95, c.a * 0.9);
    lg.ellipse('fill', x, c.y, c.w * 0.5, c.h);
    lg.ellipse('fill', x + c.w * 0.28, c.y - c.h * 0.4, c.w * 0.34, c.h * 0.8);
    lg.ellipse('fill', x - c.w * 0.28, c.y - c.h * 0.2, c.w * 0.3, c.h * 0.7);
  }
}

// the King lying FACE-UP on the carpet, drifting (drawn in screen space)
function drawLyingHero6() {
  const cx = VW * 0.5, cy = VH * 0.5 + Math.sin(T * 0.8) * 10;
  lg.push();
  lg.translate(cx, cy);
  // the magic carpet beneath him (its own hover offset is baked in)
  drawFlyingCarpet(0, 44, 2.0);
  // a simple reclining figure lying on his back, face up
  lg.push();
  lg.rotate(-Math.PI / 2);   // lay the body horizontal
  // body
  setColA(COL.shirt);
  lg.polygon('fill', -6, -26, 6, -26, 8, 20, -8, 20);
  setColA(mul(COL.vest, 0.92));
  lg.polygon('fill', -6, -20, 6, -20, 7, 6, -7, 6);
  // legs
  setColA(COL.pants || [0.2, 0.18, 0.24]); lg.setLineWidth(6);
  lg.line(-3, 18, -3, 40); lg.line(3, 18, 3, 40);
  // arms resting
  setColA(COL.skin); lg.setLineWidth(5);
  lg.line(-6, -10, -14, 6); lg.line(6, -10, 14, 6);
  // head (face up) with white hair
  setColA(COL.skin); lg.circle('fill', 0, -32, 7);
  setColA([0.9, 0.9, 0.92]); lg.circle('fill', 0, -35, 6);
  lg.pop();
  lg.pop();
  lg.setLineWidth(1);
}

// ============================================================================
//  PER-FRAME UPDATE
// ============================================================================
function updateEntsL6(dt) {
  const p = player;
  l6.msgT = Math.max(0, l6.msgT - dt);
  l6.flash = Math.max(0, l6.flash - dt);
  if (l6.quakeT > 0) { l6.quakeT -= dt; if (l6.quakeT <= 0) l6.quakeMag = 0; }
  if (l6.dialog && l6.dialog.dur > 0) l6.dialog.t += dt;
  if (l6.dialogDelay) {
    l6.dialogDelay.t += dt;
    if (l6.dialogDelay.t >= l6.dialogDelay.wait) { l6say(l6.dialogDelay.text, l6.dialogDelay.dur); l6.dialogDelay = null; }
  }
  updateL6Balls(dt);

  const w = l6.witch;

  if (l6.phase === 'summon') {
    l6.t += dt;
    w.appearT = Math.min(1, w.appearT + dt * 0.5);   // she rises slowly from the lava
    p.vx = 0; p.state = 'ground'; p.onGround = true; p.facing = -1;
    if (l6.t > 3.2 && !l6.dialogDelay && l6.sub === 0) {
      l6.sub = 1; l6say('So — the Shadow King returns. Then here you will fall for the last time.', 5);
    }
    if (l6.t > 6.4) { l6.phase = 'fight'; l6.t = 0; l6.round = 0; nextRound6(); }
    return;
  }

  if (l6.phase === 'fight') {
    // the King fights normally on the arena floor
    if (l6.vulnT > 0) l6.vulnT = Math.max(0, l6.vulnT - dt);
    w.hurtT = Math.max(0, w.hurtT - dt);
    // clamp him to the shore so he can't wander onto the lava
    if (p.x < ARENA_L6) { p.x = ARENA_L6; if (p.vx < 0) p.vx = 0; }
    if (p.x > ARENA_R6) { p.x = ARENA_R6; if (p.vx > 0) p.vx = 0; }

    if (l6.bossType === 'guardian') updateGuardian6(dt, p);
    else if (l6.bossType === 'knight') updateKnight6(dt, p);
    else { if (l6.guardian) updateGuardian6(dt, p); if (l6.knight) updateKnight6(dt, p); }

    // hero sword swing window — melee the active secondary boss
    const au = 1 - (p.atkT || 0) / ATK_DUR;
    if ((p.atkT || 0) > 0 && au > 0.30 && au < 0.56) {
      if (l6.bossType === 'guardian') tryHitGuardian6(p);
      else if (l6.bossType === 'knight') tryHitKnight6(p);
    }

    // hero lava bullets — only meaningful against the witch while she is exposed
    for (let i = l6.bullets.length - 1; i >= 0; i--) {
      const bu = l6.bullets[i];
      bu.t += dt; bu.x += bu.vx * dt; bu.y += bu.vy * dt;
      let gone = bu.t > 1.9;
      if (l6.vulnT > 0 && !w.dead && pointInWitch(bu.x, bu.y)) { hurtWitch(1); gone = true; }
      if (bu.x < ARENA_L6 - 2200) gone = true;
      if (gone) l6.bullets.splice(i, 1);
    }

    // when the current boss is gone AND the vulnerability window closes, next round
    if (!l6.bossType && !w.dead) {
      const bossGone = (!l6.guardian || l6.guardian.deadT > 1.2) && (!l6.knight || l6.knight.deadT > 1.4);
      if (l6.vulnT <= 0 && bossGone) { nextRound6(); }
    }
    return;
  }

  if (l6.phase === 'death') {
    l6.t += dt; w.deadT += dt;
    p.vx = 0; p.state = 'ground'; p.onGround = true; p.facing = -1;
    if (l6.t > 3.4 && l6.sub < 5) { l6.sub = 5; l6.phase = 'carpet'; l6.t = 0; }
    return;
  }

  if (l6.phase === 'carpet') {
    l6.t += dt;
    p.vx = 0; p.state = 'ground'; p.onGround = true; p.facing = -1;
    if (l6.t < 0.1) l6toast('The quake splits the shore — the magic carpet rises to you!');
    if (l6.t > 2.6) { startFlight6(); }
    return;
  }

  if (l6.phase === 'flight') { updateFlight6(dt); return; }
  if (l6.phase === 'castle') { updateCastle6(dt); return; }
  if (l6.phase === 'clouds') { updateClouds6(dt); return; }
}

// begin the next round of the cycle (alternating guardian / knight)
function nextRound6() {
  if (l6.witch.hp <= 0) return;
  l6.round += 1;
  if (l6.round % 2 === 1) { spawnGuardian6(); }
  else { spawnKnight6(); }
}

// ============================================================================
//  OVERLAY (HUD, witch power bar, cards, credits)
// ============================================================================
function drawL6Overlay() {
  const p = player;
  const hudOff = (l6.phase === 'summon' || l6.phase === 'death' || l6.phase === 'carpet'
    || l6.phase === 'castle' || l6.phase === 'clouds');

  if (!hudOff) {
    // hearts
    lg.setFont(FONT_HUD);
    for (let i = 1; i <= difficultyMaxHp(); i++) {
      const hx = 30 + (i - 1) * 36, hy = 32;
      const full = (p.hp || 0) >= i;
      if (full) lg.setColor(0.85, 0.16, 0.22, 1); else lg.setColor(0.25, 0.10, 0.13, 0.8);
      lg.circle('fill', hx - 5, hy - 3, 6.5); lg.circle('fill', hx + 5, hy - 3, 6.5);
      lg.polygon('fill', hx - 11, hy - 0.5, hx + 11, hy - 0.5, hx, hy + 12);
      lg.setColor(1, 1, 1, full ? 0.35 : 0.12); lg.circle('fill', hx - 6.5, hy - 5, 2);
    }
    // fire-sword charge pips
    if (p.lavaSword) {
      lg.setColor(1.0, 0.5, 0.15, 0.9); lg.print('FIRE-SWORD', 30, 60, 0, 0.85, 0.85);
      const charged = p.lavaCharge || 0;
      for (let i = 0; i < 3; i++) {
        const cx = 118 + i * 16, cy = 66;
        if (i < charged) { lg.setColor(1.0, 0.45, 0.12, 1); lg.circle('fill', cx, cy, 5); lg.setColor(1.0, 0.9, 0.5, 1); lg.circle('fill', cx - 1.4, cy - 1.4, 2); }
        else { lg.setColor(0.4, 0.2, 0.12, 0.7); lg.circle('line', cx, cy, 5); }
      }
      lg.setColor(0.85, 0.7, 0.6, 0.7);
      lg.print(l6.vulnT > 0 ? 'ATTACK: fire at the witch!' : (charged > 0 ? 'ATTACK to fire' : 'BLOCK to charge'), 178, 60, 0, 0.8, 0.8);
    }
    // WITCH power bar, top-centre
    if (l6.witch && !l6.witch.dead) {
      const wc = l6.witch;
      lg.setColor(0.7, 0.45, 0.85, 0.95);
      const gm = l6.vulnT > 0 ? 'THE  WITCH  —  EXPOSED' : 'THE  WITCH';
      lg.setFont(FONT_HUD);
      lg.print(gm, VW / 2 - FONT_HUD.getWidth(gm) / 2, 22);
      const bw = 340, bx = VW / 2 - bw / 2, by = 42;
      lg.setColor(0.12, 0.06, 0.16, 0.85); lg.rectangle('fill', bx, by, bw, 10);
      lg.setColor(l6.vulnT > 0 ? 0.95 : 0.6, l6.vulnT > 0 ? 0.4 : 0.85, l6.vulnT > 0 ? 0.3 : 0.95, 1);
      lg.rectangle('fill', bx, by, bw * clamp(wc.hp / wc.maxHp, 0, 1), 10);
      lg.setColor(1, 0.9, 1, 0.5); lg.rectangle('fill', bx, by, bw, 2);
    }
    // secondary boss "blows remaining"
    if (l6.bossType === 'guardian' && l6.guardian && !l6.guardian.dead) {
      const b = l6.guardian;
      lg.setColor(0.9, 0.3, 0.25, 0.9);
      const gm = 'GUARDIAN';
      lg.print(gm, VW / 2 - FONT_HUD.getWidth(gm) / 2, 66);
      const bw = 220, bx = VW / 2 - bw / 2, by = 84;
      lg.setColor(0.2, 0.06, 0.06, 0.8); lg.rectangle('fill', bx, by, bw, 8);
      lg.setColor(0.85, 0.20, 0.18, 1); lg.rectangle('fill', bx, by, bw * clamp(b.hp / 6, 0, 1), 8);
    } else if (l6.bossType === 'knight' && l6.knight && !l6.knight.dead) {
      const k = l6.knight;
      lg.setColor(0.95, 0.4, 0.2, 0.9);
      const gm = 'LAVA  KNIGHT';
      lg.print(gm, VW / 2 - FONT_HUD.getWidth(gm) / 2, 66);
      const bw = 220, bx = VW / 2 - bw / 2, by = 84;
      lg.setColor(0.2, 0.06, 0.04, 0.8); lg.rectangle('fill', bx, by, bw, 8);
      lg.setColor(0.95, 0.35, 0.12, 1); lg.rectangle('fill', bx, by, bw * clamp(k.hp / 4, 0, 1), 8);
    }
    // flight progress
    if (l6.flight && l6.flight.active && (l6.flight.phase === 'run' || l6.flight.phase === 'lift')) {
      const f = l6.flight;
      const prog = clamp((f.startX - p.x) / (f.startX - f.holeX), 0, 1);
      const bw = 300, bx = VW / 2 - bw / 2, by = 26;
      lg.setColor(0.2, 0.06, 0.04, 0.7); lg.rectangle('fill', bx, by, bw, 8);
      lg.setColor(0.8, 0.75, 0.5, 1); lg.rectangle('fill', bx, by, bw * prog, 8);
      lg.setColor(0.9, 0.85, 0.75, 0.9);
      const m = 'BACK  ACROSS  THE  LAVA  RIVER';
      lg.print(m, VW / 2 - FONT_HUD.getWidth(m) / 2, 40);
    }
  }

  // red flash on the witch being wounded / her death
  if (l6.flash > 0) {
    lg.setColor(1.0, 0.4, 0.3, clamp(l6.flash / 0.7, 0, 1) * 0.6);
    lg.rectangle('fill', 0, 0, VW, VH);
  }

  // toast + dialog
  if (l6.msgT > 0) {
    lg.setFont(FONT_HUD);
    lg.setColor(0.96, 0.88, 0.78, Math.min(1, l6.msgT));
    lg.print(l6.msg, VW / 2 - FONT_HUD.getWidth(l6.msg) / 2, VH - 96);
  }
  if (l6.dialog && l6.dialog.dur > 0 && l6.dialog.t < l6.dialog.dur) {
    drawSubtitle({ who: l6.phase === 'summon' || l6.phase === 'fight' && l6.sub === 1 ? 'WITCH' : 'HERO', text: l6.dialog.text });
  }

  // summon location card
  if (l6.phase === 'summon' && FONT_LOC) {
    const a = clamp((l6.t - 0.4) / 1.0, 0, 1) * clamp((6.0 - l6.t) / 1.2, 0, 1);
    if (a > 0) {
      lg.setFont(FONT_LOC);
      lg.setColor(0.9, 0.7, 0.95, a);
      printSpaced("THE  WITCH'S  RECKONING", VW / 2, VH * 0.16, FONT_LOC, 5, 1);
    }
  }

  // castle: fade IN from black (arriving through the hole), quaking
  if (l6.phase === 'castle') {
    const a = clamp(1 - l6.t / 1.2, 0, 1);
    if (a > 0) { lg.setColor(0, 0, 0, a); lg.rectangle('fill', 0, 0, VW, VH); }
    const b = clamp((l6.t - 4.6) / 0.9, 0, 1);   // fade out to the sky
    if (b > 0) { lg.setColor(0, 0, 0, b); lg.rectangle('fill', 0, 0, VW, VH); }
    if (FONT_SUB) {
      const ta = clamp((l6.t - 1.2) / 1.0, 0, 1) * clamp((4.4 - l6.t) / 1.0, 0, 1);
      if (ta > 0) {
        lg.setFont(FONT_SUB); lg.setColor(0.9, 0.85, 0.8, ta);
        printSpaced('OUT  OF  THE  BURNING  KEEP', VW / 2, VH * 0.2, FONT_SUB, 4, 0.9);
      }
    }
  }

  // clouds + credits
  if (l6.phase === 'clouds') {
    const fadeIn = clamp(l6.t / 1.4, 0, 1);
    if (fadeIn < 1) { lg.setColor(0, 0, 0, 1 - fadeIn); lg.rectangle('fill', 0, 0, VW, VH); }
    if (FONT_SUB && FONT_TITLE) {
      if (l6.credits >= 1) {
        const a = clamp((l6.end.t - 3.0) / 1.2, 0, 1) * (l6.credits >= 2 ? clamp((6.5 - l6.end.t) / 1.0, 0, 1) : 1);
        lg.setFont(FONT_SUB); lg.setColor(0.15, 0.13, 0.16, a);
        printSpaced('Directed  and  programmed  by  Francesco  Nicolosi', VW / 2, VH * 0.42, FONT_SUB, 2, 0.7);
      }
      if (l6.credits >= 2) {
        const a = clamp((l6.end.t - 6.5) / 1.2, 0, 1) * (l6.credits >= 3 ? clamp((10.0 - l6.end.t) / 1.0, 0, 1) : 1);
        lg.setFont(FONT_SUB); lg.setColor(0.15, 0.13, 0.16, a);
        printSpaced('Nycosoft  presented', VW / 2, VH * 0.42, FONT_SUB, 4, 0.9);
      }
      if (l6.credits >= 3) {
        const a = clamp((l6.end.t - 10.0) / 1.6, 0, 1);
        lg.setFont(FONT_TITLE);
        const offs = [[-2, 0], [2, 0], [0, -2], [0, 2], [0, 0]];
        for (const off of offs) {
          if (off[0] === 0 && off[1] === 0) setColA(COL.title, a);
          else lg.setColor(1, 0.85, 0.55, a * 0.10);
          printSpaced('THE RETURN OF THE SHADOW', VW / 2 + off[0], VH * 0.36 + off[1], FONT_TITLE, 13, 0.9);
        }
        lg.setFont(FONT_HUD);
        lg.setColor(0.2, 0.18, 0.22, a * 0.7);
        const m = 'press  R  to  play  again';
        lg.print(m, VW / 2 - FONT_HUD.getWidth(m) / 2, VH - 60);
      }
    }
  }
}
