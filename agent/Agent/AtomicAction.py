from abc import ABC, abstractmethod
from typing import List, Dict

class Action(ABC):
    '''
    The Action interface declares atomic actions common to all supported backends.
    '''
    _validActions = {
        'CopyPaste': {
            'call': None,
            'description': None,
            'parameters': None
        },
        'AutoFill': {
            'call': None,
            'description': None,
            'parameters': None
        },
        'Write': {
            'call': None,
            'description': None,
            'parameters': None
        },
    }

    @property
    def validActions(self) -> dict:
        return self._validActions
                
    @abstractmethod
    def Write(self) -> None:
        '''
        Writes a value to a Range.
        '''
        pass

    @abstractmethod
    def CopyPaste(self,) -> None:
        '''
        Copies the value and format of a source Range to a destination Range.
        '''
        pass
    
    @abstractmethod
    def CopyPasteFormat(self,) -> None:
        '''
        Copies the format of a source Range to a destination Range.
        '''
        pass

    @abstractmethod
    def AutoFill(self,) -> None:
        '''
        Autofills the value of a source Range to a destination Range.
        '''
        pass

    @abstractmethod
    def Sort(self,) -> None:
        '''
        Sort the source Range.
        '''
        pass

    @abstractmethod
    def Filter(self,) -> None:
        '''
        Filter the source Range based on the key1 Range.
        '''
        pass

    @abstractmethod
    def DeleteFilter(self,) -> None:
        """
        Delete all filters.
        """
        pass
    
    @abstractmethod
    def RemoveDuplicate(self,) -> None:
        """
        Removes duplicate values from a range of values.
        """
        pass

    @abstractmethod
    def group_ungroup(self) -> None:
        pass

    @abstractmethod
    def HideRow(self,) -> None:
        """
        Hide one row.
        """
        pass
    
    @abstractmethod
    def HideColumn(self,) -> None:
        """
        Hide one column.
        """
        pass
    
    @abstractmethod
    def UnhideRow(self,) -> None:
        """
        Show one hidden row.
        """
        pass
    
    @abstractmethod
    def UnhideColumn(self,) -> None:
        """
        Show one hidden column.
        """
        pass

    @abstractmethod
    def Merge(self,) -> None:
        """
        Creates a merged cell from the specified Range object.
        """
        pass

    @abstractmethod
    def Unmerge(self,) -> None:
        """
        Separates a merged area into individual cells.
        """
        pass

    def merging_text(self) -> None:
        pass

    @abstractmethod
    def Delete(self,) -> None:
        """
        Deletes a cell or range of cells.
        """
        pass

    @abstractmethod
    def Clear(self,) -> None:
        """
        Clears the content and the formatting of a Range.
        """
        pass

    @abstractmethod
    def InsertRow(self,) -> None:
        """
        Insert one row.
        """
        pass
    
    @abstractmethod
    def InsertColumn(self,) -> None:
        """
        Insert one column.
        """
        pass

    @abstractmethod
    def SplitText(self,) -> None:
        """
        Splite text in one column to multiple column.
        """
        pass

    @abstractmethod
    def AutoFit(self,) -> None:
        """
        Autofits the width and height of all cells in the range.
        """
        pass