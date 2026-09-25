const { test } = require('node:test');
const assert = require('node:assert/strict');
const { zoneAt, goalTargetAt, automaticShotResult, availableThrowResults,
  hasShotOrigin, hasGoalPoint, hasShotPoint, pointShots, oldShotGames, computeHeat,
  throwSituationOf, filterThrowGames } = require('../.shot-analysis-test.cjs');

test('field and goal taps resolve precise positions and legacy categories', () => {
  assert.equal(zoneAt(200, 140), 'DURCH_M');
  assert.equal(zoneAt(200, 40), 'KREIS');
  assert.equal(goalTargetAt(50, 50), 't1');
  assert.equal(goalTargetAt(200, 120), 't5');
  assert.equal(goalTargetAt(200, 40), 'POST');
  assert.equal(goalTargetAt(15, 120), 'WIDE');
});

test('goal position limits the possible throw results', () => {
  const shot = { type: 'throw', shotX: 0.5, shotY: 0.5, goalX: 0.5, goalY: 0.5 };
  assert.equal(automaticShotResult(goalTargetAt(200, 40)), 'post');
  assert.equal(automaticShotResult(goalTargetAt(15, 120)), 'wide');
  assert.equal(automaticShotResult(goalTargetAt(200, 120)), null);
  assert.deepEqual(availableThrowResults({ ...shot, target: 'POST' }), ['post']);
  assert.deepEqual(availableThrowResults({ ...shot, target: 'WIDE' }), ['wide']);
  assert.deepEqual(availableThrowResults({ ...shot, target: 't5' }), ['goal', 'saved']);
  assert.deepEqual(availableThrowResults({ type: 'throw', target: 'POST' }), ['goal', 'saved', 'post', 'wide']);
});

test('a blocked throw has an origin, counts as a shot, and has no goal target', () => {
  const blocked = { id: 'blocked', type: 'throw', side: 'us', playerId: 'p1', zone: 'RUECK_M',
    result: 'blocked', target: null, shotX: 0.5, shotY: 0.8, goalX: null, goalY: null };
  const game = { actions: [blocked] };
  assert.equal(hasShotOrigin(blocked), true);
  assert.equal(hasGoalPoint(blocked), false);
  assert.equal(hasShotPoint(blocked), false);
  assert.deepEqual(availableThrowResults(blocked), ['blocked']);
  assert.equal(pointShots([game], 'p1', false).length, 1);
  assert.equal(computeHeat([game], 'p1', false).total.n, 1);
  assert.equal(computeHeat(oldShotGames([game]), 'p1', false).total.n, 0);
  assert.equal(computeHeat([game], 'p1', false).targets[null], undefined);
});

test('situation filters preserve exact zones while selecting the throw context', () => {
  const counter = { id: 'counter', type: 'throw', side: 'us', playerId: 'p1', zone: 'RUECK_L', situation: 'konter', result: 'goal' };
  const normal = { id: 'normal', type: 'throw', side: 'us', playerId: 'p1', zone: 'RUECK_L', result: 'goal' };
  const game = { actions: [counter, normal] };
  assert.equal(throwSituationOf(counter), 'konter');
  assert.equal(throwSituationOf(normal), 'normal');
  assert.deepEqual(filterThrowGames([game], 'konter')[0].actions.map((a) => a.id), ['counter']);
  assert.equal(filterThrowGames([game], 'konter')[0].actions[0].zone, 'RUECK_L');
});

test('precise shots retain results and old throws remain available', () => {
  const fresh = { id: 'new', type: 'throw', side: 'us', playerId: 'p1', zone: 'KREIS',
    target: 't5', result: 'goal', shotX: 0, shotY: 0.25, goalX: 0.5, goalY: 0.5 };
  const old = { id: 'old', type: 'throw', side: 'us', playerId: 'p1', zone: 'RUECK_M',
    target: 'POST', result: 'post' };
  const keeper = { ...fresh, id: 'keeper', side: 'them', keeperId: 'tw', result: 'saved' };
  const game = { actions: [fresh, old, keeper] };
  assert.equal(hasShotPoint(fresh), true); // 0 is a valid coordinate
  assert.equal(hasShotPoint(old), false);
  assert.deepEqual(pointShots([game], 'p1', false).map((x) => x.id), ['new']);
  assert.deepEqual(pointShots([game], 'tw', true).map((x) => x.id), ['keeper']);
  assert.equal(computeHeat(oldShotGames([game]), 'p1', false).total.n, 1);
  assert.equal(computeHeat([game], 'p1', false).total.n, 2);
});
