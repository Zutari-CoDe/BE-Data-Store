# GitHub Pages Deployment Guide

This file documents how the BE-Data-Store React app was hosted on GitHub Pages.

## 1. Configure `package.json`

1. Add the `homepage` field with the GitHub Pages URL:
   ```json
   "homepage": "https://zutari-code.github.io/BE-Data-Store"
   ```
2. Add deployment scripts:
   ```json
   "scripts": {
     "start": "react-scripts start",
     "build": "react-scripts build",
     "predeploy": "npm run build",
     "deploy": "gh-pages -d build",
     "test": "react-scripts test",
     "eject": "react-scripts eject"
   }
   ```

## 2. Install `gh-pages`

Run:
```bash
npm install --save-dev gh-pages
```

This package publishes the build output to the `gh-pages` branch automatically.

## 3. Deploy the app

Run:
```bash
npm run deploy
```

This command performs:
- `npm run build`
- `gh-pages -d build`

After completion, the app is published to:

`https://zutari-code.github.io/BE-Data-Store/`

## 4. GitHub repository settings

For a public repo, the GitHub Pages site is usually served from the `gh-pages` branch.

No additional manual repo settings were required because the deployment script created the branch automatically.

## 5. Notes

- Use the `homepage` field when your app is not deployed at the root of a domain.
- If you change the repo name or username, update `homepage` accordingly.
- Re-run `npm run deploy` after future changes.
