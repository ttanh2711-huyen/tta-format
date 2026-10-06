/** Các phép thử chạy trong lượt "Chạy tất cả": bỏ những phép thử đánh dấu `risky` (chạy riêng). */
export function probesForRunAll<T extends { risky?: boolean }>(probes: readonly T[]): T[] {
  return probes.filter((probe) => !probe.risky);
}
