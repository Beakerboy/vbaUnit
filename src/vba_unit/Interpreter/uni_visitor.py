from pyvba_interpreter.vba_listener import VbaListener
from typing import TypeVar


T = TypeVar('T', bound='VbaUnitListener')


class VbaUnitListener(VbaListener):
    pass
