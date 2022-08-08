#!/usr/bin/env python
"""Django's command-line utility for administrative tasks."""
import os
import sys


def main():
    print("Please Wait Scyil pvt Ltd server is starting on /Dettol Software /user 09253 ")
    print("Please Wait Scyil pvt Ltd server is starting on /Dettol Software /user 09255 ")
    os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'dettol.settings')
    try:
        from django.core.management import execute_from_command_line
    except ImportError as exc:
        raise ImportError(
            "Couldn't import Django. Are you sure it's installed and "
            "available on your PYTHONPATH environment variable? Did you "
            "forget to activate a virtual environment?"
        ) from exc
    execute_from_command_line(sys.argv)

if __name__ == '__main__':
    main()