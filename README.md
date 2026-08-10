# PLRA → PERA DC Valuation Rate API Specification

Formal Government-to-Government (G2G) API Specification Document for exposing PLRA DC Valuation Rate services to the PERA Activity Register.

## Deliverables

| File | Description |
|------|-------------|
| `PLRA_PERA_DC_Valuation_API_Specification.docx` | Full API specification (template-aligned) |
| `generate_pera_api_spec.py` | Regenerates the Word document |

## Document identity

- **Identifier:** `PLRA-DCVAL-API-SPEC-001`
- **Version:** `0.2` (Draft)
- **Provider:** Punjab Land Records Authority (PLRA)
- **Consumer:** PERA — Activity Register Application
- **Classification:** Restricted

## API catalogue (v1)

| API ID | Operation | Path focus |
|--------|-----------|------------|
| API-01 | Get Property Types | Rural / Urban |
| API-02 | Get All Districts | District list |
| API-03 | Get Tehsils by District | Cascading tehsil |
| API-04 | Get Mauzas by Tehsil | Urban path |
| API-05 | Get Khasra Nos by Mauza | Urban (incl. Square/Kila ids) |
| API-06 | Get Property Classifications | Rural path |
| API-07 | Get Property Areas by Tehsil | Rural path |
| API-08 | Get Khasra / Kila Nos by Property Area | Rural multi-select |
| API-09 | Get Valuation by Khasra No (**GetValuationMethod**) | Urban rate |
| API-10 | Get Valuation by Property Area | Rural full-area rate |
| API-11 | Get Valuation by Property Area & Khasra No(s) (**GetValuationMethod**) | Rural multi-khasra rates |

## Consumer journeys

**Urban:** Property Type → District → Tehsil → Mauza → Khasra → Valuation (API-09 with khasraId + squareId + kilaId)

**Rural (full area):** Property Type → District → Tehsil → Classification → Property Area → Valuation (API-10)

**Rural (khasra select):** Property Type → District → Tehsil → Classification → Property Area → Khasra(s) → Valuation (API-11)

## Valuation response fields

`dcRate`, `structureRate`, `location`, `classification`, `unitOfMeasurement`

## Regenerate

```bash
pip install python-docx
python3 generate_pera_api_spec.py
```
