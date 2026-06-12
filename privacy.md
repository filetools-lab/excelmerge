# Privacy Notes

Excel Report Merger is designed for local-first processing.

## What stays local

- Uploaded workbooks are read in the browser.
- Header checks, mismatch detection, merging, previews, and exports run on the user's device.
- Workbook contents are not uploaded to an application server.
- The app does not store selected files after the page is closed.

## What may use the network

- The current web version loads SheetJS from the official SheetJS CDN.
- This network request downloads the library code; it does not upload workbook content.

## Future paid version

If payment or licensing is added later, the recommended design is:

- Keep workbook processing local.
- Use a third-party payment provider for checkout.
- Use a small license-check service only for license status.
- Do not send file names, sheet names, headers, rows, or workbook contents to the license service.
