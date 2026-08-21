// 「現在可以吃什麼」功能單元測試（零食／早餐類別 + 晚餐熱量預留）
// 執行方式： node --test tests/
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { loadApp } = require('./load-app.js');

const t = loadApp(path.join(__dirname, '..', 'index.html'));

function goals(overrides){
  return Object.assign({ protein: 35, fat: 35, sugar: 31, cal: 1250 }, overrides || {});
}

test('snackCategoryFoods 只挑出類別是「零食」或「早餐」的食材，其他類別不會出現', () => {
  const foods = [
    { id: 'a', name: '洋芋片', category: '零食' },
    { id: 'b', name: '吐司', category: '早餐' },
    { id: 'c', name: '雞胸肉', category: '肉類' },
    { id: 'd', name: '沒分類的東西', category: '' },
  ];
  const result = t.snackCategoryFoods(foods);
  const names = result.map((f) => f.name);
  assert.deepEqual(names.sort(), ['吐司', '洋芋片']);
});

test('computeSnackBudget 用晚餐預留熱量的「上限」（較保守）去扣除，確保晚餐一定還有預算', () => {
  const budget = t.computeSnackBudget({ protein: 0, fat: 0, sugar: 0, cal: 300 }, goals(), 500, 700);
  // 今天目標1250，已吃300，剩950；扣掉晚餐上限700 => 現在可用 250
  assert.equal(budget.calLeftTotal, 950);
  assert.equal(budget.calAvailableNow, 250);
});

test('computeSnackBudget 如果使用者把晚餐預留下限/上限填反，仍會自動修正（取較大值當上限）', () => {
  const budget = t.computeSnackBudget({ protein: 0, fat: 0, sugar: 0, cal: 300 }, goals(), 700, 500);
  assert.equal(budget.dinnerMin, 500);
  assert.equal(budget.dinnerMax, 700);
  assert.equal(budget.calAvailableNow, 250);
});

test('computeSnackBudget 熱量已經吃超過「目標減晚餐預留」時，現在可用熱量會是負數（代表不建議再吃）', () => {
  const budget = t.computeSnackBudget({ protein: 0, fat: 0, sugar: 0, cal: 900 }, goals(), 500, 700);
  // 剩 350，扣掉晚餐 700 => -350
  assert.equal(budget.calAvailableNow, -350);
});

test('computeSnackBudget 正確計算脂肪與糖的剩餘額度', () => {
  const budget = t.computeSnackBudget({ protein: 0, fat: 20, sugar: 25, cal: 300 }, goals(), 500, 700);
  assert.equal(budget.fatLeft, 15); // 35 - 20
  assert.equal(budget.sugarLeft, 6); // 31 - 25
});

test('maxGramsForFood 以「熱量／脂肪／糖」三者中最嚴格的限制為準', () => {
  // 熱量限制換算：可吃 250kcal / (400kcal/100g) *100 = 62.5g
  // 脂肪限制換算：可吃 15g / (20g/100g) *100 = 75g
  // 糖限制換算：  可吃 6g / (40g/100g) *100 = 15g   <-- 這個最嚴格
  const budget = t.computeSnackBudget({ protein: 0, fat: 20, sugar: 25, cal: 300 }, goals(), 500, 700);
  const food = { base: 100, servings: 5, protein100: 5, fat100: 20, sugar100: 40, cal100: 400 };
  const maxG = t.maxGramsForFood(food, budget);
  assert.equal(Math.floor(maxG), 15);
});

test('maxGramsForFood 對 base 不是 100g 的食材（例如一份 30g 的餅乾）要用 base 換算，不能誤當成「每100g」（迴歸測試）', () => {
  // 這款餅乾一份是 30g，「每 30g」480kcal / 20g脂肪 / 10g糖。
  // 熱量限制：550kcal 可用額度 / 480kcal（每30g）* 30g ≈ 34.375g
  // 脂肪限制：35g 可用額度 / 20g（每30g）* 30g = 52.5g
  // 糖限制：  31g 可用額度 / 10g（每30g）* 30g = 93g
  // 庫存上限：30g * 3份 = 90g
  // 最嚴格的是熱量限制 ≈34.375g，floor 後是 34g。
  const budget = t.computeSnackBudget({ protein: 0, fat: 0, sugar: 0, cal: 0 }, goals(), 500, 700);
  const food = { base: 30, servings: 3, protein100: 8, fat100: 20, sugar100: 10, cal100: 480 };
  const maxG = t.maxGramsForFood(food, budget);
  assert.equal(Math.floor(maxG), 34);
  // 用 computeMacros 反向驗證：吃 maxG 克時的熱量不應該超過可用額度太多
  const m = t.computeMacros(food, Math.floor(maxG));
  assert.ok(m.cal <= budget.calAvailableNow + 1); // 容許 1kcal 的四捨五入誤差
});

test('maxGramsForFood 現在可用熱量為 0 或負數時，回傳 0（不建議吃）', () => {
  const budget = t.computeSnackBudget({ protein: 0, fat: 0, sugar: 0, cal: 900 }, goals(), 500, 700);
  const food = { base: 100, servings: 1, protein100: 5, fat100: 5, sugar100: 5, cal100: 200 };
  assert.equal(t.maxGramsForFood(food, budget), 0);
});

test('maxGramsForFood 對熱量/脂肪/糖都是 0 的食材（例如無糖氣泡水），不會被熱量或巨量營養素限制，但仍以食材庫存總重量為上限', () => {
  const budget = t.computeSnackBudget({ protein: 0, fat: 0, sugar: 0, cal: 300 }, goals(), 500, 700);
  const food = { base: 500, servings: 1, protein100: 0, fat100: 0, sugar100: 0, cal100: 0 };
  const maxG = t.maxGramsForFood(food, budget);
  assert.equal(maxG, 500); // 天花板 = base * servings
});

test('maxGramsForFood 就算熱量／脂肪／糖額度都很寬鬆，建議份量也不會超過食材本身的總庫存量（一份克數 × 份數）', () => {
  const budget = t.computeSnackBudget({ protein: 0, fat: 0, sugar: 0, cal: 0 }, goals({ cal: 5000, fat: 500, sugar: 500 }), 0, 0);
  const food = { base: 20, servings: 3, protein100: 1, fat100: 1, sugar100: 1, cal100: 10 }; // 庫存只有 60g
  const maxG = t.maxGramsForFood(food, budget);
  assert.equal(maxG, 60);
});

test('已知脂肪額度用完（fatLeft<=0）時，含脂肪的食材可吃量會被限制為 0，但完全不含脂肪的食材不受影響', () => {
  const budget = t.computeSnackBudget({ protein: 0, fat: 35, sugar: 0, cal: 300 }, goals(), 500, 700);
  assert.equal(budget.fatLeft, 0);
  const fattyFood = { base: 100, servings: 1, protein100: 5, fat100: 10, sugar100: 0, cal100: 150 };
  const fatFreeFood = { base: 100, servings: 1, protein100: 5, fat100: 0, sugar100: 5, cal100: 100 };
  assert.equal(t.maxGramsForFood(fattyFood, budget), 0);
  assert.ok(t.maxGramsForFood(fatFreeFood, budget) > 0);
});
