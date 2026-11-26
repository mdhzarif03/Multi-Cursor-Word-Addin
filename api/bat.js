// Serves a tiny Windows installer/uninstaller with this site's address baked in.
module.exports = (req, res) => {
  const host = req.headers['x-forwarded-host'] || req.headers.host;
  const site = 'https://' + host;
  const uninstall = req.query && req.query.mode === 'uninstall';
  const flag = uninstall ? ' -Uninstall' : '';
  const lines = [
    '@echo off',
    'title Multi-Cursor for Word - ' + (uninstall ? 'Uninstall' : 'Install'),
    'powershell -NoProfile -ExecutionPolicy Bypass -Command "$p=Join-Path $env:TEMP \'mc-setup.ps1\'; Invoke-WebRequest \'' + site + '/install.ps1\' -UseBasicParsing -OutFile $p; & $p -Site \'' + site + '\'' + flag + '"',
    'echo.',
    'pause'
  ];
  res.setHeader('Content-Type', 'application/octet-stream');
  res.setHeader('Content-Disposition', 'attachment; filename="' + (uninstall ? 'Uninstall-MultiCursor.bat' : 'Install-MultiCursor.bat') + '"');
  res.setHeader('Cache-Control', 'no-store');
  res.status(200).send(lines.join('\r\n') + '\r\n');
};
