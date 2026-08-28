/* Informational counts only: never added to the material bill or exports. */
function insulatorRate(head) {
  const name = String(head || "").trim().toUpperCase();
  if (/^CCB\s*,\s*CCB$/.test(name)) return [6, 0];
  // Special phases, doubled heads and combined assemblies need their own rates.
  if (/^\d|1\s*-?\s*P|[,+]/.test(name)) return null;
  if (/^CCB(?:\s|$)/.test(name)) return [name.includes("ประกบ") ? 6 : 3, 0];
  const rates = { "DDE.BL": [0, 24], DDE: [6, 24], DE: [0, 12], BA: [4, 12], SP: [3, 0], DP: [6, 0] };
  for (const [kind, rate] of Object.entries(rates)) {
    if (name === kind || name.startsWith(kind + ".") || name.startsWith(kind + " ")) return rate;
  }
  return null;
}

function insulatorQuantity(value) {
  const source = String(value ?? "").trim();
  if (!source || source.length > 100) return null;
  const tokens = source.match(/(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?|[()+-]|\S/g) || [];
  let index = 0;
  function atom() {
    const token = tokens[index++];
    if (token === "+") return atom();
    if (token === "-") return -atom();
    if (token === "(") {
      const result = sum();
      if (tokens[index++] !== ")") throw new Error("Unclosed parenthesis");
      return result;
    }
    if (!token || !/^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(token)) throw new Error("Invalid number");
    return Number(token);
  }
  function sum() {
    let result = atom();
    while (tokens[index] === "+" || tokens[index] === "-") {
      const operator = tokens[index++];
      const right = atom();
      result = operator === "+" ? result + right : result - right;
    }
    return result;
  }
  try {
    const result = sum();
    return index === tokens.length && Number.isFinite(result) ? result : null;
  } catch { return null; }
}

function summarizeInsulators(pages) {
  const result = { upright: 0, horizontal: 0, warnings: [] };
  pages.forEach((page, pageIndex) => page.forEach((row, rowIndex) => {
    if (!row.head || !row.size) return;
    const quantity = insulatorQuantity(row.count);
    const label = `หน้า ${pageIndex + 1} แถว ${rowIndex + 1}: ${row.head}`;
    if (quantity === null || quantity < 0) {
      result.warnings.push(`${label} — กรุณาตรวจจำนวน`);
      return;
    }
    if (quantity === 0) return;
    const rate = insulatorRate(row.head);
    if (!rate) {
      result.warnings.push(`${label} — ยังไม่มีอัตราลูกถ้วย`);
      return;
    }
    result.upright += rate[0] * quantity;
    result.horizontal += rate[1] * quantity;
  }));
  return result;
}

if (typeof module !== "undefined") module.exports = { insulatorRate, insulatorQuantity, summarizeInsulators };
