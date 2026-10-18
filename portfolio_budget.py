"""Whole-share allocation with hard percentage caps and fees inside each cap.
Unused allocations/rounding cash are preserved, never silently redistributed.
"""
from decimal import Decimal, ROUND_FLOOR, localcontext


def allocate_shares(prices, budget, fee, weights=None):
    with localcontext() as context:
        # Covers the entire finite IEEE float range without rounding budget invariants.
        context.prec = 700
        prices = [Decimal(str(p)) for p in prices]
        budget, fee = Decimal(str(budget)), Decimal(str(fee))
        if not budget.is_finite() or not fee.is_finite() or budget <= 0 or fee < 0:
            raise ValueError('Invalid budget or fee')
        if not prices or any(not p.is_finite() or p <= 0 for p in prices):
            raise ValueError('Prices must be finite and positive')
        if any(abs(v.adjusted()) > 308 for v in [budget, fee, *prices] if v):
            raise ValueError('Numeric magnitude exceeds supported range')
        n = len(prices)
        if weights is None:
            weights = [Decimal(100) / n] * n
        else:
            weights = [Decimal(str(w)) for w in weights]
        if len(weights) != n or any(not w.is_finite() or w < 0 for w in weights):
            raise ValueError('Invalid allocation weights')
        total_weight = sum(weights)
        if total_weight > Decimal('100.00000001'):
            raise ValueError('Allocations exceed 100 percent')
        # Normalize only floating-point epsilon above 100%; never invest excess budget.
        if total_weight > 100:
            weights = [w * 100 / total_weight for w in weights]
        counts = [max(0, int(((budget * w / 100 - fee) / p).to_integral_value(rounding=ROUND_FLOOR)))
                  for p, w in zip(prices, weights)]
        investments = [p * c + (fee if c else 0) for p, c in zip(prices, counts)]
        if sum(investments) > budget:
            raise ValueError('Allocation exceeds budget')
        return counts, [float(v) for v in investments]
