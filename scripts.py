import logging
from pathlib import Path
from garminexport.incremental_backup import incremental_backup
from garminexport.logging_config import LOG_LEVELS

logging.basicConfig(
    level=logging.INFO, format="%(asctime)-15s [%(levelname)s] %(message)s"
)
log = logging.getLogger(__name__)


def export_fit_files():
    username = "athoren79"
    password = None
    backup_dir = Path("D:\\Andreas\\OneDrive\\Andreas\\träning\\Running\\Workouts")
    auth_token_dir = Path().home() / ".garminexport"
    log_level = "INFO"
    export_format = ["fit"]
    ignore_errors = False
    max_retries = 3

    logging.root.setLevel(LOG_LEVELS[log_level])

    try:
        incremental_backup(
            username=username,
            password=password,
            auth_token_dir=auth_token_dir,
            backup_dir=backup_dir,
            export_formats=export_format,
            ignore_errors=ignore_errors,
            max_retries=max_retries,
        )
    except Exception as e:
        log.error(str(e))


if __name__ == "__main__":
    export_fit_files()
