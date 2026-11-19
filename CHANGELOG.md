# Changelog - gcv/updater (GCV Updater add-in)

## 2.0.0

- Separate Excel add-in "GCV Updater.xlam" with the ribbon button "update tools".
- Reads the catalog (catalog.json in Gaza-children-village/excel-updater) and compares every installed GCV add-in with its newest release; installs updates and new required add-ins after one question, offers new optional add-ins one by one.
- Installs while Excel runs: the old file is kept as a fallback and put back if the new version does not start; that version is not offered again. Its own update is installed by GCV Loader at the next Excel start.
- Checks size and SHA-256 of every download. No access to the VBA project is needed.
