
from abc import ABC, abstractmethod
from enum import Enum, unique, auto


class Switch(ABC):

    @unique
    class State(Enum):
        ON = auto()
        OFF = auto()
        NA = auto()

    @property
    @abstractmethod
    def state(self):
        pass

    @abstractmethod
    def turn_on(self):
        pass

    @abstractmethod
    def turn_off(self):
        pass
