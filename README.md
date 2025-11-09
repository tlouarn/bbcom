# bbcom
Excel VBA wrapper for the low-level Bloomberg COM API. Currently supports low-level calls to `BDP` (Bloomberg Data Point) and `BDH` (Bloomberg Data History) calls.

## Installation
1. Open a new Excel workbook
2. Create a new Class Module named `BBCOM` and paste the code from the file `bbcom.vba`
3. Create a new Module and paste the test functions in the file `test.vba`
4. Run the functions
5. Save as a macro-enable Excel workbook ('.xlsm' or '.xlsb')

## Unadjusted historical prices
I have added support for unadjusted historical prices based on the BLPAPI 3.5 COM Developer Guide (link below).

The four arguments are:
- `adjustmentFollowDPDF`
- `adjustmentAbnormal`
- `adjustmentSplit`
- `adjustmentNormal`

The arguments work as follow:
- if `adjustmentFollowDPDF` is set to `True` then the request will ignore the other arguments and follow the user's DPDF settings
- if `adjustmentFollowDPDF` is set to `False` then the request will take the other arguments into account. If nothing is specified then no adjustment is made.

## Troubleshoot

Make sure the Trust Center settings allow macros to run

File > Options > Trust Center > Trust Center Settings > Macro Settings > Enable All MAcros

<img width="826" height="678" alt="image" src="https://github.com/user-attachments/assets/af2b579d-84df-498a-bf77-496e43aad9fd" />

## References
* [Original code](https://github.com/conzchung/Bloomberg-VBA/blob/main/Code)
* [BLPAPI 3.5 COM Developer Guide](https://data.bloomberglp.com/professional/sites/10/2017/03/BLPAPI-Core-Developer-Guide.pdf)
