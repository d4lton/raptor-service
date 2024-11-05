#
# Copyright ©2024 Dana Basken
#

from dynaconf import Dynaconf

settings = Dynaconf(
    envvar_prefix="RAPTOR",
    settings_files=["config.json", ".secrets.json"]
)