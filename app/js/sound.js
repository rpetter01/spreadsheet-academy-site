/* Spreadsheet Academy sound: tiny Web Audio synth, no audio files. window.Sound */
(function (root) {
'use strict';
const KEY = 'gridwise.sound.v1';
let ctx = null, master = null, last = {};
let on = true;
try { on = localStorage.getItem(KEY) !== 'off'; } catch (e) { /* storage unavailable */ }

function ensure() {
  if (ctx) { if (ctx.state === 'suspended') ctx.resume(); return ctx; }
  const AC = root.AudioContext || root.webkitAudioContext; if (!AC) return null;
  ctx = new AC(); master = ctx.createGain(); master.gain.value = 0.9; master.connect(ctx.destination); return ctx;
}
/** One enveloped note. */
function tone(freq, t0, dur, { type = 'sine', gain = 0.18, slideTo = null, attack = 0.008 } = {}) {
  const o = ctx.createOscillator(), g = ctx.createGain(), t = ctx.currentTime + t0;
  o.type = type; o.frequency.setValueAtTime(freq, t);
  if (slideTo) o.frequency.exponentialRampToValueAtTime(slideTo, t + dur);
  g.gain.setValueAtTime(0.0001, t); g.gain.exponentialRampToValueAtTime(gain, t + attack); g.gain.exponentialRampToValueAtTime(0.0001, t + dur);
  o.connect(g); g.connect(master); o.start(t); o.stop(t + dur + 0.03);
}
const N = { C5: 523.25, D5: 587.33, E5: 659.25, G5: 783.99, A5: 880, B5: 987.77, C6: 1046.5, E6: 1318.5, G4: 392, A4: 440 };
const SOUNDS = {
  correct() { tone(N.C5, 0, .16, { type: 'triangle', gain: .2 }); tone(N.E5, .08, .16, { type: 'triangle', gain: .2 }); tone(N.G5, .16, .34, { type: 'triangle', gain: .22 }); tone(N.C6, .16, .34, { type: 'sine', gain: .08 }); },
  wrong() { tone(220, 0, .16, { type: 'square', gain: .07, slideTo: 180 }); tone(165, .13, .26, { type: 'square', gain: .07, slideTo: 130 }); },
  complete() { [N.C5, N.E5, N.G5, N.C6].forEach((f, i) => tone(f, i * .11, .3, { type: 'triangle', gain: .2 })); tone(N.E6, .46, .7, { type: 'sine', gain: .12 }); tone(N.G5, .46, .7, { type: 'triangle', gain: .14 }); tone(N.C6, .46, .7, { type: 'triangle', gain: .14 }); },
  next() { tone(N.A4, 0, .07, { type: 'sine', gain: .1 }); tone(N.E5, .05, .1, { type: 'sine', gain: .1 }); },
  hint() { tone(N.G5, 0, .18, { type: 'sine', gain: .1 }); tone(N.B5, .1, .22, { type: 'sine', gain: .08 }); },
  enter() { tone(660, 0, .05, { type: 'triangle', gain: .07 }); },
  type() { tone(1500 + Math.random() * 300, 0, .02, { type: 'square', gain: .012, attack: .002 }); },
  blocked() { tone(150, 0, .12, { type: 'sine', gain: .12, slideTo: 110 }); },
};
const Sound = {
  get on() { return on; },
  play(name) {
    if (!on || !SOUNDS[name]) return;
    const now = performance.now(); if (last[name] && now - last[name] < 25) return; last[name] = now;
    try { if (ensure()) SOUNDS[name](); } catch (e) { /* audio unavailable */ }
  },
  toggle() {
    on = !on; try { localStorage.setItem(KEY, on ? 'on' : 'off'); } catch (e) { /* ignore */ }
    if (on) this.play('enter'); return on;
  },
};
root.Sound = Sound;
})(window);
