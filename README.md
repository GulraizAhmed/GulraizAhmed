# PLRA Central Payment Gateway — New Client Onboarding Guide

Formal Standard Operating Procedure (SOP) and integration guide for onboarding a new client department / service onto the PLRA Central Payment Gateway.

## Deliverables

| File | Description |
|------|-------------|
| `PLRA_Payment_Gateway_New_Client_Onboarding_Guide.docx` | Full onboarding SOP + sample payloads |
| `generate_payment_gateway_onboarding.py` | Regenerates the Word document |

## Document identity

- **Identifier:** `PLRA-PGW-ONB-GUIDE-001`
- **Version:** `1.1`
- **Classification:** Restricted
- **Audience:** New client departments/services that own challans/fees

## What the guide covers

1. Payment Gateway solution overview (Web Portal, Punjab Zameen, Admin Portal)
2. Scope — CLRMIS & E-Stamp today; new services via prefix registry
3. **SOP** for onboarding a new client (10 steps)
4. Runtime flow: Token → Fetch → (MPGS pay) → Intimate
5. **Pattern A** authentication (mandatory for new services)
6. Sample **request/response payloads** for:
   - Token API
   - Fetch (unpaid / already-paid / not-found)
   - Intimate (success / already-paid / idempotent retry)
7. Dual-side network whitelisting (PLRA ↔ Client)
8. Pre-go-live information exchange checklist
9. **UAT checklist**
10. **Live / Production checklist**
11. Operational expectations, FAQ, onboarding message template, sign-off

## Key rule for new clients

You expose **exactly three HTTP APIs**. PLRA Gateway calls you. You do **not** call MPGS or BoP for this flow.

## Regenerate

```bash
pip install python-docx
python3 generate_payment_gateway_onboarding.py
```
