# Deployment Notes

This directory keeps host- and service-level deployment helpers that belong to
the operational side of this repository.

## Included scripts

- `update.sh`
  Updates the checked-out branch, recreates `venv/` if missing, refreshes
  Python dependencies, fixes runtime directory ownership, and restarts the
  systemd service.
- `sync_nginx_container_certs.sh`
  Copies Let's Encrypt certificates into the host paths expected by the nginx
  container, then reloads nginx inside the container.

## Safety notes

- `update.sh` uses `git reset --hard origin/<branch>`.
  This will discard uncommitted changes in the deployment checkout.
- `update.sh` assumes the production service runs as `www-data` and uses
  `/tmp/excel_webdatamerger` plus `/opt/excel_webdatamerger/logs` as writable
  runtime directories.
- Both scripts assume a specific server layout.
  Review paths, service names, domain names, and container names before use.
- Run these scripts only on the intended server.

## Suggested usage

```bash
chmod +x deploy/update.sh deploy/sync_nginx_container_certs.sh
./deploy/update.sh
```
