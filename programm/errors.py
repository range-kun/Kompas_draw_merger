from __future__ import annotations


class ExecutionNotSelectedError(Exception):
    pass


class SpecificationEmptyError(Exception):
    pass


class DifferentDrawsForSameOboznError(Exception):
    pass


class NoDrawsError(Exception):
    pass


class FolderNotSelectedError(Exception):
    pass
