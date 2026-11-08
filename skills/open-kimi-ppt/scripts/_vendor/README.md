# Vendored pure-Python dependencies

Copied unmodified so the export scripts run without `pip install` (and so without
network access) on embedded Pythons such as the one bundled with desktop agents.
They are used only when the module is not already importable.

| Package | Version | License | Source |
|---|---|---|---|
| PyYAML (`yaml/`, pure-Python part only; the optional libyaml extension is omitted) | 6.0.3 | MIT (`yaml/LICENSE`) | https://pypi.org/project/PyYAML/ |
