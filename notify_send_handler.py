#!/usr/bin/env python

"""
Provides a logging handler that emits messages using notify-send.

These messages are usually displayed as transient popup windows on the
desktop.

To use, add code like this to main():

    # Also log using notify-send if it is available.
    if notify_send_handler.NotifySendHandler.is_available():
        logger.addHandler(
            notify_send_handler.NotifySendHandler(
                os.path.splitext(os.path.basename(__file__))[0]
            )
        )
"""

import logging
import subprocess


class NotifySendHandler(logging.Handler):
    """A logging Handler that displays records using notify-send."""
    levelname_to_icon = {
        "DEBUG": None,
        "INFO": "dialog-information",
        "WARNING": "dialog-warning",
        "ERROR": "dialog-error",
        "CRITICAL": "dialog-error"
    }

    def __init__(self, summary=None, expire_time=None):
        logging.Handler.__init__(self)
        self.summary = summary
        self.expire_time = expire_time
        self.setFormatter(logging.Formatter('%(msg)s'))

    def emit(self, record):
        """Override the default handler's emit method."""
        message = self.format(record)
        args = ['notify-send']
        if self.expire_time:
            args.append("--expire-time=%s" % self.expire_time)
        icon = NotifySendHandler.levelname_to_icon.get(record.levelname, None)
        if icon:
            args.append("--icon=%s" % icon)
        if self.summary:
            args.append(self.summary)
        args.append(message)
        subprocess.check_call(args)

    @staticmethod
    def is_available():
        """Returns True if the notify-send program is available."""
        p = subprocess.Popen(['which', 'notify-send'], stdout=subprocess.PIPE)
        p.communicate()
        return p.returncode == 0
