# GEU AQAR Consolidator — Streamlit v1

This converts the working AQAR scanner into a browser UI.

## Local use — recommended for the 7 GB OneDrive repository

```bash
python3 -m pip install -r requirements.txt
streamlit run app.py
```

Then enter the local OneDrive AQAR root, e.g.

`/Users/yourname/OneDrive/AQAR`

The files remain on the local computer. They are not uploaded by the app in Local mode.

## ZIP mode

Useful for small departmental tests such as HSS. Do not upload the 7 GB repository to a hosted Streamlit server.

## Important

The app distinguishes:
- department folders
- Criteria-* folders
- Data Template information
- Information Not in Data Template

The non-template requirements are built from the supplied departmental requirements document:
1.1.1, 1.3.1, 2.2.1, 2.3.1, 2.3.2, 2.3.3, 2.6.1, 2.6.2, 2.7.1, 4.1.1, 6.5.2, 6.5.3.

All proposed institutional values remain `PENDING IQAC APPROVAL`.
