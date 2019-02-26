#!/usr/bin/env python
"""
ViperMonkey: Execution context for global and local variables

ViperMonkey is a specialized engine to parse, analyze and interpret Microsoft
VBA macros (Visual Basic for Applications), mainly for malware analysis.

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

Project Repository:
https://github.com/decalage2/ViperMonkey
"""

# === LICENSE ==================================================================

# ViperMonkey is copyright (c) 2015-2016 Philippe Lagadec (http://www.decalage.info)
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without modification,
# are permitted provided that the following conditions are met:
#
#  * Redistributions of source code must retain the above copyright notice, this
#    list of conditions and the following disclaimer.
#  * Redistributions in binary form must reproduce the above copyright notice,
#    this list of conditions and the following disclaimer in the documentation
#    and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
# ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
# FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
# DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
# SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
# CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
# OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
# OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

__version__ = '0.02'

# --- IMPORTS ------------------------------------------------------------------

import xlrd

import array
import os
from hashlib import sha256
from datetime import datetime
from .logger import log
import base64

def is_procedure(vba_object):
    """
    Check if a VBA object is a procedure, e.g. a Sub or a Function.
    This is implemented by checking if the object has a statements
    attribute
    :param vba_object: VBA_Object to be checked
    :return: True if vba_object is a procedure, False otherwise
    """
    if hasattr(vba_object, 'statements'):
        return True
    else:
        return False

# === VBA CLASSES =====================================================================================================

# global dictionary of constants, functions and subs for the VBA library
VBA_LIBRARY = {}

# Output directory to save dropped artifacts.
out_dir = None
# Count of files dropped.
file_count = 0

class Context(object):
    """
    a Context object contains the global and local named objects (variables, subs, functions)
    used to evaluate VBA statements.
    """

    def __init__(self,
                 _globals=None,
                 _locals=None,
                 context=None,
                 engine=None,
                 doc_vars=None,
                 loaded_excel=None):

        # Track mapping from bogus alias name of DLL imported functions to
        # real names.
        self.dll_func_true_names = {}
        
        # Track a dict mapping the labels of code blocks labeled with the LABEL:
        # construct to code blocks. This will be used to evaluate GOTO statements
        # when emulating.
        self.tagged_blocks = {}

        # Track the in-memory loaded Excel workbook (xlrd workbook object).
        self.loaded_excel = loaded_excel
        
        # Track open files.
        self.open_files = {}

        # Track the final contents of written files.
        self.closed_files = {}

        # globals should be a pointer to the globals dict from the core VBA engine (ViperMonkey)
        # because each statement should be able to change global variables
        if _globals is not None:
            # direct copy of the pointer to globals:
            self.globals = _globals
        elif context is not None:
            self.globals = context.globals
            self.open_files = context.open_files
            self.closed_files = context.closed_files
            self.loaded_excel = context.loaded_excel
            self.dll_func_true_names = context.dll_func_true_names
        else:
            self.globals = {}
        # on the other hand, each Context should have its own private copy of locals
        if _locals is not None:
            # However, if locals is explicitly provided, we use a copy of it:
            self.locals = dict(_locals)
        else:
            self.locals = {}
        # engine should be a pointer to the ViperMonkey engine, to provide callback features
        if engine is not None:
            self.engine = engine
        elif context is not None:
            self.engine = context.engine
        else:
            self.engine = None

        log.debug("Have xlrd loaded Excel file = " + str(self.loaded_excel is not None))
            
        # Track data saved in document variables.
        if doc_vars is not None:
            # direct copy of the pointer to globals:
            self.doc_vars = doc_vars
        elif context is not None:
            self.doc_vars = context.doc_vars
        else:
            self.doc_vars = {}
            
        # Track whether nested loops are running with a stack of flags. If a loop is
        # running its flag will be True.
        self.loop_stack = []

        # Track whether we have exited from the current function.
        self.exit_func = False

        # Track variable types, if known.
        self.types = {}

        # Track the current with prefix for with statements.
        self.with_prefix = ""

        # Add in a global for the current time.
        self.globals["Now".lower()] = datetime.now()

        # Fake up a user name.
        self.globals["Application.UserName".lower()] = "--"
        
        # Add some attributes we are handling as global variables.
        self.globals["vbDirectory".lower()] = "vbDirectory"
        self.globals["VBA.vbDirectory".lower()] = "vbDirectory"
        self.globals["VBA.KeyCodeConstants.vbDirectory".lower()] = "vbDirectory"
        self.globals["vbKeyLButton".lower()] = 1
        self.globals["VBA.vbKeyLButton".lower()] = 1
        self.globals["VBA.KeyCodeConstants.vbKeyLButton".lower()] = 1
        self.globals["vbKeyRButton".lower()] = 2
        self.globals["VBA.vbKeyRButton".lower()] = 2
        self.globals["VBA.KeyCodeConstants.vbKeyRButton".lower()] = 2
        self.globals["vbKeyCancel".lower()] = 3
        self.globals["VBA.vbKeyCancel".lower()] = 3
        self.globals["VBA.KeyCodeConstants.vbKeyCancel".lower()] = 3
        self.globals["vbKeyMButton".lower()] = 4
        self.globals["VBA.vbKeyMButton".lower()] = 4
        self.globals["VBA.KeyCodeConstants.vbKeyMButton".lower()] = 4
        self.globals["vbKeyBack".lower()] = 8
        self.globals["VBA.vbKeyBack".lower()] = 8
        self.globals["VBA.KeyCodeConstants.vbKeyBack".lower()] = 8
        self.globals["vbKeyTab".lower()] = 9
        self.globals["VBA.vbKeyTab".lower()] = 9
        self.globals["VBA.KeyCodeConstants.vbKeyTab".lower()] = 9
        self.globals["vbKeyClear".lower()] = 12
        self.globals["VBA.vbKeyClear".lower()] = 12
        self.globals["VBA.KeyCodeConstants.vbKeyClear".lower()] = 12
        self.globals["vbKeyReturn".lower()] = 13
        self.globals["VBA.vbKeyReturn".lower()] = 13
        self.globals["VBA.KeyCodeConstants.vbKeyReturn".lower()] = 13
        self.globals["vbKeyShift".lower()] = 16
        self.globals["VBA.vbKeyShift".lower()] = 16
        self.globals["VBA.KeyCodeConstants.vbKeyShift".lower()] = 16
        self.globals["vbKeyControl".lower()] = 17
        self.globals["VBA.vbKeyControl".lower()] = 17
        self.globals["VBA.KeyCodeConstants.vbKeyControl".lower()] = 17
        self.globals["vbKeyMenu".lower()] = 18
        self.globals["VBA.vbKeyMenu".lower()] = 18
        self.globals["VBA.KeyCodeConstants.vbKeyMenu".lower()] = 18
        self.globals["vbKeyPause".lower()] = 19
        self.globals["VBA.vbKeyPause".lower()] = 19
        self.globals["VBA.KeyCodeConstants.vbKeyPause".lower()] = 19
        self.globals["vbKeyCapital".lower()] = 20
        self.globals["VBA.vbKeyCapital".lower()] = 20
        self.globals["VBA.KeyCodeConstants.vbKeyCapital".lower()] = 20
        self.globals["vbKeyEscape".lower()] = 27
        self.globals["VBA.vbKeyEscape".lower()] = 27
        self.globals["VBA.KeyCodeConstants.vbKeyEscape".lower()] = 27
        self.globals["vbKeySpace".lower()] = 32
        self.globals["VBA.vbKeySpace".lower()] = 32
        self.globals["VBA.KeyCodeConstants.vbKeySpace".lower()] = 32
        self.globals["vbKeyPageUp".lower()] = 33
        self.globals["VBA.vbKeyPageUp".lower()] = 33
        self.globals["VBA.KeyCodeConstants.vbKeyPageUp".lower()] = 33
        self.globals["vbKeyPageDown".lower()] = 34
        self.globals["VBA.vbKeyPageDown".lower()] = 34
        self.globals["VBA.KeyCodeConstants.vbKeyPageDown".lower()] = 34
        self.globals["vbKeyEnd".lower()] = 35
        self.globals["VBA.vbKeyEnd".lower()] = 35
        self.globals["VBA.KeyCodeConstants.vbKeyEnd".lower()] = 35
        self.globals["vbKeyHome".lower()] = 36
        self.globals["VBA.vbKeyHome".lower()] = 36
        self.globals["VBA.KeyCodeConstants.vbKeyHome".lower()] = 36
        self.globals["vbKeyLeft".lower()] = 37
        self.globals["VBA.vbKeyLeft".lower()] = 37
        self.globals["VBA.KeyCodeConstants.vbKeyLeft".lower()] = 37
        self.globals["vbKeyUp".lower()] = 38
        self.globals["VBA.vbKeyUp".lower()] = 38
        self.globals["VBA.KeyCodeConstants.vbKeyUp".lower()] = 38
        self.globals["vbKeyRight".lower()] = 39
        self.globals["VBA.vbKeyRight".lower()] = 39
        self.globals["VBA.KeyCodeConstants.vbKeyRight".lower()] = 39
        self.globals["vbKeyDown".lower()] = 40
        self.globals["VBA.vbKeyDown".lower()] = 40
        self.globals["VBA.KeyCodeConstants.vbKeyDown".lower()] = 40
        self.globals["vbKeySelect".lower()] = 41
        self.globals["VBA.vbKeySelect".lower()] = 41
        self.globals["VBA.KeyCodeConstants.vbKeySelect".lower()] = 41
        self.globals["vbKeyPrint".lower()] = 42
        self.globals["VBA.vbKeyPrint".lower()] = 42
        self.globals["VBA.KeyCodeConstants.vbKeyPrint".lower()] = 42
        self.globals["vbKeyExecute".lower()] = 43
        self.globals["VBA.vbKeyExecute".lower()] = 43
        self.globals["VBA.KeyCodeConstants.vbKeyExecute".lower()] = 43
        self.globals["vbKeySnapshot".lower()] = 44
        self.globals["VBA.vbKeySnapshot".lower()] = 44
        self.globals["VBA.KeyCodeConstants.vbKeySnapshot".lower()] = 44
        self.globals["vbKeyInsert".lower()] = 45
        self.globals["VBA.vbKeyInsert".lower()] = 45
        self.globals["VBA.KeyCodeConstants.vbKeyInsert".lower()] = 45
        self.globals["vbKeyDelete".lower()] = 46
        self.globals["VBA.vbKeyDelete".lower()] = 46
        self.globals["VBA.KeyCodeConstants.vbKeyDelete".lower()] = 46
        self.globals["vbKeyHelp".lower()] = 47
        self.globals["VBA.vbKeyHelp".lower()] = 47
        self.globals["VBA.KeyCodeConstants.vbKeyHelp".lower()] = 47
        self.globals["vbKeyNumlock".lower()] = 144
        self.globals["VBA.vbKeyNumlock".lower()] = 144
        self.globals["VBA.KeyCodeConstants.vbKeyNumlock".lower()] = 144        
        self.globals["vbKeyA".lower()] = 65
        self.globals["VBA.vbKeyA".lower()] = 65
        self.globals["VBA.KeyCodeConstants.vbKeyA".lower()] = 65
        self.globals["vbKeyB".lower()] = 66
        self.globals["VBA.vbKeyB".lower()] = 66
        self.globals["VBA.KeyCodeConstants.vbKeyB".lower()] = 66
        self.globals["vbKeyC".lower()] = 67
        self.globals["VBA.vbKeyC".lower()] = 67
        self.globals["VBA.KeyCodeConstants.vbKeyC".lower()] = 67
        self.globals["vbKeyD".lower()] = 68
        self.globals["VBA.vbKeyD".lower()] = 68
        self.globals["VBA.KeyCodeConstants.vbKeyD".lower()] = 68
        self.globals["vbKeyE".lower()] = 69
        self.globals["VBA.vbKeyE".lower()] = 69
        self.globals["VBA.KeyCodeConstants.vbKeyE".lower()] = 69
        self.globals["vbKeyF".lower()] = 70
        self.globals["VBA.vbKeyF".lower()] = 70
        self.globals["VBA.KeyCodeConstants.vbKeyF".lower()] = 70
        self.globals["vbKeyG".lower()] = 71
        self.globals["VBA.vbKeyG".lower()] = 71
        self.globals["VBA.KeyCodeConstants.vbKeyG".lower()] = 71
        self.globals["vbKeyH".lower()] = 72
        self.globals["VBA.vbKeyH".lower()] = 72
        self.globals["VBA.KeyCodeConstants.vbKeyH".lower()] = 72
        self.globals["vbKeyI".lower()] = 73
        self.globals["VBA.vbKeyI".lower()] = 73
        self.globals["VBA.KeyCodeConstants.vbKeyI".lower()] = 73
        self.globals["vbKeyJ".lower()] = 74
        self.globals["VBA.vbKeyJ".lower()] = 74
        self.globals["VBA.KeyCodeConstants.vbKeyJ".lower()] = 74
        self.globals["vbKeyK".lower()] = 75
        self.globals["VBA.vbKeyK".lower()] = 75
        self.globals["VBA.KeyCodeConstants.vbKeyK".lower()] = 75
        self.globals["vbKeyL".lower()] = 76
        self.globals["VBA.vbKeyL".lower()] = 76
        self.globals["VBA.KeyCodeConstants.vbKeyL".lower()] = 76
        self.globals["vbKeyM".lower()] = 77
        self.globals["VBA.vbKeyM".lower()] = 77
        self.globals["VBA.KeyCodeConstants.vbKeyM".lower()] = 77
        self.globals["vbKeyN".lower()] = 78
        self.globals["VBA.vbKeyN".lower()] = 78
        self.globals["VBA.KeyCodeConstants.vbKeyN".lower()] = 78
        self.globals["vbKeyO".lower()] = 79
        self.globals["VBA.vbKeyO".lower()] = 79
        self.globals["VBA.KeyCodeConstants.vbKeyO".lower()] = 79
        self.globals["vbKeyP".lower()] = 80
        self.globals["VBA.vbKeyP".lower()] = 80
        self.globals["VBA.KeyCodeConstants.vbKeyP".lower()] = 80
        self.globals["vbKeyQ".lower()] = 81
        self.globals["VBA.vbKeyQ".lower()] = 81
        self.globals["VBA.KeyCodeConstants.vbKeyQ".lower()] = 81
        self.globals["vbKeyR".lower()] = 82
        self.globals["VBA.vbKeyR".lower()] = 82
        self.globals["VBA.KeyCodeConstants.vbKeyR".lower()] = 82
        self.globals["vbKeyS".lower()] = 83
        self.globals["VBA.vbKeyS".lower()] = 83
        self.globals["VBA.KeyCodeConstants.vbKeyS".lower()] = 83
        self.globals["vbKeyT".lower()] = 84
        self.globals["VBA.vbKeyT".lower()] = 84
        self.globals["VBA.KeyCodeConstants.vbKeyT".lower()] = 84
        self.globals["vbKeyU".lower()] = 85
        self.globals["VBA.vbKeyU".lower()] = 85
        self.globals["VBA.KeyCodeConstants.vbKeyU".lower()] = 85
        self.globals["vbKeyV".lower()] = 86
        self.globals["VBA.vbKeyV".lower()] = 86
        self.globals["VBA.KeyCodeConstants.vbKeyV".lower()] = 86
        self.globals["vbKeyW".lower()] = 87
        self.globals["VBA.vbKeyW".lower()] = 87
        self.globals["VBA.KeyCodeConstants.vbKeyW".lower()] = 87
        self.globals["vbKeyX".lower()] = 88
        self.globals["VBA.vbKeyX".lower()] = 88
        self.globals["VBA.KeyCodeConstants.vbKeyX".lower()] = 88
        self.globals["vbKeyY".lower()] = 89
        self.globals["VBA.vbKeyY".lower()] = 89
        self.globals["VBA.KeyCodeConstants.vbKeyY".lower()] = 89
        self.globals["vbKeyZ".lower()] = 90
        self.globals["VBA.vbKeyZ".lower()] = 90
        self.globals["VBA.KeyCodeConstants.vbKeyZ".lower()] = 90
        self.globals["vbKey0".lower()] = 48
        self.globals["VBA.vbKey0".lower()] = 48
        self.globals["VBA.KeyCodeConstants.vbKey0".lower()] = 48
        self.globals["vbKey1".lower()] = 49
        self.globals["VBA.vbKey1".lower()] = 49
        self.globals["VBA.KeyCodeConstants.vbKey1".lower()] = 49
        self.globals["vbKey2".lower()] = 50
        self.globals["VBA.vbKey2".lower()] = 50
        self.globals["VBA.KeyCodeConstants.vbKey2".lower()] = 50
        self.globals["vbKey3".lower()] = 51
        self.globals["VBA.vbKey3".lower()] = 51
        self.globals["VBA.KeyCodeConstants.vbKey3".lower()] = 51
        self.globals["vbKey4".lower()] = 52
        self.globals["VBA.vbKey4".lower()] = 52
        self.globals["VBA.KeyCodeConstants.vbKey4".lower()] = 52
        self.globals["vbKey5".lower()] = 53
        self.globals["VBA.vbKey5".lower()] = 53
        self.globals["VBA.KeyCodeConstants.vbKey5".lower()] = 53
        self.globals["vbKey6".lower()] = 54
        self.globals["VBA.vbKey6".lower()] = 54
        self.globals["VBA.KeyCodeConstants.vbKey6".lower()] = 54
        self.globals["vbKey7".lower()] = 55
        self.globals["VBA.vbKey7".lower()] = 55
        self.globals["VBA.KeyCodeConstants.vbKey7".lower()] = 55
        self.globals["vbKey8".lower()] = 56
        self.globals["VBA.vbKey8".lower()] = 56
        self.globals["VBA.KeyCodeConstants.vbKey8".lower()] = 56
        self.globals["vbKey9".lower()] = 57
        self.globals["VBA.vbKey9".lower()] = 57
        self.globals["VBA.KeyCodeConstants.vbKey9".lower()] = 57
        self.globals["vbKeyNumpad0".lower()] = 96
        self.globals["VBA.vbKeyNumpad0".lower()] = 96
        self.globals["VBA.KeyCodeConstants.vbKeyNumpad0".lower()] = 96
        self.globals["vbKeyNumpad1".lower()] = 97
        self.globals["VBA.vbKeyNumpad1".lower()] = 97
        self.globals["VBA.KeyCodeConstants.vbKeyNumpad1".lower()] = 97
        self.globals["vbKeyNumpad2".lower()] = 98
        self.globals["VBA.vbKeyNumpad2".lower()] = 98
        self.globals["VBA.KeyCodeConstants.vbKeyNumpad2".lower()] = 98
        self.globals["vbKeyNumpad3".lower()] = 99
        self.globals["VBA.vbKeyNumpad3".lower()] = 99
        self.globals["VBA.KeyCodeConstants.vbKeyNumpad3".lower()] = 99
        self.globals["vbKeyNumpad4".lower()] = 100
        self.globals["VBA.vbKeyNumpad4".lower()] = 100
        self.globals["VBA.KeyCodeConstants.vbKeyNumpad4".lower()] = 100
        self.globals["vbKeyNumpad5".lower()] = 101
        self.globals["VBA.vbKeyNumpad5".lower()] = 101
        self.globals["VBA.KeyCodeConstants.vbKeyNumpad5".lower()] = 101
        self.globals["vbKeyNumpad6".lower()] = 102
        self.globals["VBA.vbKeyNumpad6".lower()] = 102
        self.globals["VBA.KeyCodeConstants.vbKeyNumpad6".lower()] = 102
        self.globals["vbKeyNumpad7".lower()] = 103
        self.globals["VBA.vbKeyNumpad7".lower()] = 103
        self.globals["VBA.KeyCodeConstants.vbKeyNumpad7".lower()] = 103
        self.globals["vbKeyNumpad8".lower()] = 104
        self.globals["VBA.vbKeyNumpad8".lower()] = 104
        self.globals["VBA.KeyCodeConstants.vbKeyNumpad8".lower()] = 104
        self.globals["vbKeyNumpad9".lower()] = 105
        self.globals["VBA.vbKeyNumpad9".lower()] = 105
        self.globals["VBA.KeyCodeConstants.vbKeyNumpad9".lower()] = 105
        self.globals["vbKeyMultiply".lower()] = 106
        self.globals["VBA.vbKeyMultiply".lower()] = 106
        self.globals["VBA.KeyCodeConstants.vbKeyMultiply".lower()] = 106
        self.globals["vbKeyAdd".lower()] = 107
        self.globals["VBA.vbKeyAdd".lower()] = 107
        self.globals["VBA.KeyCodeConstants.vbKeyAdd".lower()] = 107
        self.globals["vbKeySeparator".lower()] = 108
        self.globals["VBA.vbKeySeparator".lower()] = 108
        self.globals["VBA.KeyCodeConstants.vbKeySeparator".lower()] = 108
        self.globals["vbKeySubtract".lower()] = 109
        self.globals["VBA.vbKeySubtract".lower()] = 109
        self.globals["VBA.KeyCodeConstants.vbKeySubtract".lower()] = 109
        self.globals["vbKeyDecimal".lower()] = 110
        self.globals["VBA.vbKeyDecimal".lower()] = 110
        self.globals["VBA.KeyCodeConstants.vbKeyDecimal".lower()] = 110
        self.globals["vbKeyDivide".lower()] = 111
        self.globals["VBA.vbKeyDivide".lower()] = 111
        self.globals["VBA.KeyCodeConstants.vbKeyDivide".lower()] = 111
        self.globals["vbKeyF1".lower()] = 112
        self.globals["VBA.vbKeyF1".lower()] = 112
        self.globals["VBA.KeyCodeConstants.vbKeyF1".lower()] = 112
        self.globals["vbKeyF2".lower()] = 113
        self.globals["VBA.vbKeyF2".lower()] = 113
        self.globals["VBA.KeyCodeConstants.vbKeyF2".lower()] = 113
        self.globals["vbKeyF3".lower()] = 114
        self.globals["VBA.vbKeyF3".lower()] = 114
        self.globals["VBA.KeyCodeConstants.vbKeyF3".lower()] = 114
        self.globals["vbKeyF4".lower()] = 115
        self.globals["VBA.vbKeyF4".lower()] = 115
        self.globals["VBA.KeyCodeConstants.vbKeyF4".lower()] = 115
        self.globals["vbKeyF5".lower()] = 116
        self.globals["VBA.vbKeyF5".lower()] = 116
        self.globals["VBA.KeyCodeConstants.vbKeyF5".lower()] = 116
        self.globals["vbKeyF6".lower()] = 117
        self.globals["VBA.vbKeyF6".lower()] = 117
        self.globals["VBA.KeyCodeConstants.vbKeyF6".lower()] = 117
        self.globals["vbKeyF7".lower()] = 118
        self.globals["VBA.vbKeyF7".lower()] = 118
        self.globals["VBA.KeyCodeConstants.vbKeyF7".lower()] = 118
        self.globals["vbKeyF8".lower()] = 119
        self.globals["VBA.vbKeyF8".lower()] = 119
        self.globals["VBA.KeyCodeConstants.vbKeyF8".lower()] = 119
        self.globals["vbKeyF9".lower()] = 120
        self.globals["VBA.vbKeyF9".lower()] = 120
        self.globals["VBA.KeyCodeConstants.vbKeyF9".lower()] = 120
        self.globals["vbKeyF10".lower()] = 121
        self.globals["VBA.vbKeyF10".lower()] = 121
        self.globals["VBA.KeyCodeConstants.vbKeyF10".lower()] = 121
        self.globals["vbKeyF11".lower()] = 122
        self.globals["VBA.vbKeyF11".lower()] = 122
        self.globals["VBA.KeyCodeConstants.vbKeyF11".lower()] = 122
        self.globals["vbKeyF12".lower()] = 123
        self.globals["VBA.vbKeyF12".lower()] = 123
        self.globals["VBA.KeyCodeConstants.vbKeyF12".lower()] = 123
        self.globals["vbKeyF13".lower()] = 124
        self.globals["VBA.vbKeyF13".lower()] = 124
        self.globals["VBA.KeyCodeConstants.vbKeyF13".lower()] = 124
        self.globals["vbKeyF14".lower()] = 125
        self.globals["VBA.vbKeyF14".lower()] = 125
        self.globals["VBA.KeyCodeConstants.vbKeyF14".lower()] = 125
        self.globals["vbKeyF15".lower()] = 126
        self.globals["VBA.vbKeyF15".lower()] = 126
        self.globals["VBA.KeyCodeConstants.vbKeyF15".lower()] = 126
        self.globals["vbKeyF16".lower()] = 127
        self.globals["VBA.vbKeyF16".lower()] = 127
        self.globals["VBA.KeyCodeConstants.vbKeyF16".lower()] = 127
        self.globals["vbNullString".lower()] = ''
        self.globals["VBA.vbNullString".lower()] = ''
        self.globals["VBA.KeyCodeConstants.vbNullString".lower()] = ''
        self.globals["vbNullChar".lower()] = '\0'
        self.globals["VBA.vbNullChar".lower()] = '\0'
        self.globals["VBA.KeyCodeConstants.vbNullChar".lower()] = '\0'

        self.globals["vbUpperCase".lower()] = 1
        self.globals["VBA.vbUpperCase".lower()] = 1
        self.globals["VBA.KeyCodeConstants.vbUpperCase".lower()] = 1
        self.globals["vbLowerCase".lower()] = 2
        self.globals["VBA.vbLowerCase".lower()] = 2
        self.globals["VBA.KeyCodeConstants.vbLowerCase".lower()] = 2
        self.globals["vbProperCase".lower()] = 3
        self.globals["VBA.vbProperCase".lower()] = 3
        self.globals["VBA.KeyCodeConstants.vbProperCase".lower()] = 3
        self.globals["vbWide".lower()] = 4
        self.globals["VBA.vbWide".lower()] = 4
        self.globals["VBA.KeyCodeConstants.vbWide".lower()] = 4
        self.globals["vbNarrow".lower()] = 8
        self.globals["VBA.vbNarrow".lower()] = 8
        self.globals["VBA.KeyCodeConstants.vbNarrow".lower()] = 8
        self.globals["vbKatakana".lower()] = 16
        self.globals["VBA.vbKatakana".lower()] = 16
        self.globals["VBA.KeyCodeConstants.vbKatakana".lower()] = 16
        self.globals["vbHiragana".lower()] = 32
        self.globals["VBA.vbHiragana".lower()] = 32
        self.globals["VBA.KeyCodeConstants.vbHiragana".lower()] = 32
        self.globals["vbUnicode".lower()] = 64
        self.globals["VBA.vbUnicode".lower()] = 64
        self.globals["VBA.KeyCodeConstants.vbUnicode".lower()] = 64
        self.globals["vbFromUnicode".lower()] = 128
        self.globals["VBA.vbFromUnicode".lower()] = 128
        self.globals["VBA.KeyCodeConstants.vbFromUnicode".lower()] = 128

        self.globals["xlOuterCenterPoint".lower()] = 2.0
        self.globals["xlPivotLineBlank".lower()] = 2
        self.globals["rgbMaroon".lower()] = 128

        self.globals["vbKeyLButton".lower()] = 0x1
        self.globals["vbKeyRButton".lower()] = 0x2
        self.globals["vbKeyCancel".lower()] = 0x3
        self.globals["vbKeyMButton".lower()] = 0x4
        self.globals["vbKeyBack".lower()] = 0x8
        self.globals["vbKeyTab".lower()] = 0x9
        self.globals["vbKeyClear".lower()] = 0xC
        self.globals["vbKeyReturn".lower()] = 0xD
        self.globals["vbKeyShift".lower()] = 0x10
        self.globals["vbKeyControl".lower()] = 0x11
        self.globals["vbKeyMenu".lower()] = 0x12
        self.globals["vbKeyPause".lower()] = 0x13
        self.globals["vbKeyCapital".lower()] = 0x14
        self.globals["vbKeyEscape".lower()] = 0x1B
        self.globals["vbKeySpace".lower()] = 0x20
        self.globals["vbKeyPageUp".lower()] = 0x21
        self.globals["vbKeyPageDown".lower()] = 0x22
        self.globals["vbKeyEnd".lower()] = 0x23
        self.globals["vbKeyHome".lower()] = 0x24
        self.globals["vbKeyLeft".lower()] = 0x25
        self.globals["vbKeyUp".lower()] = 0x26
        self.globals["vbKeyRight".lower()] = 0x27
        self.globals["vbKeyDown".lower()] = 0x28
        self.globals["vbKeySelect".lower()] = 0x29
        self.globals["vbKeyPrint".lower()] = 0x2A
        self.globals["vbKeyExecute".lower()] = 0x2B
        self.globals["vbKeySnapshot".lower()] = 0x2C
        self.globals["vbKeyInsert".lower()] = 0x2D
        self.globals["vbKeyDelete".lower()] = 0x2E
        self.globals["vbKeyHelp".lower()] = 0x2F
        self.globals["vbKeyNumlock".lower()] = 0x90

        self.globals["vbKeyA".lower()] = 65
        self.globals["vbKeyB".lower()] = 66
        self.globals["vbKeyC".lower()] = 67
        self.globals["vbKeyD".lower()] = 68
        self.globals["vbKeyE".lower()] = 69
        self.globals["vbKeyF".lower()] = 70
        self.globals["vbKeyG".lower()] = 71
        self.globals["vbKeyH".lower()] = 72
        self.globals["vbKeyI".lower()] = 73
        self.globals["vbKeyJ".lower()] = 74
        self.globals["vbKeyK".lower()] = 75
        self.globals["vbKeyL".lower()] = 76
        self.globals["vbKeyM".lower()] = 77
        self.globals["vbKeyN".lower()] = 78
        self.globals["vbKeyO".lower()] = 79
        self.globals["vbKeyP".lower()] = 80
        self.globals["vbKeyQ".lower()] = 81
        self.globals["vbKeyR".lower()] = 82
        self.globals["vbKeyS".lower()] = 83
        self.globals["vbKeyT".lower()] = 84
        self.globals["vbKeyU".lower()] = 85
        self.globals["vbKeyV".lower()] = 86
        self.globals["vbKeyW".lower()] = 87
        self.globals["vbKeyX".lower()] = 88
        self.globals["vbKeyY".lower()] = 89
        self.globals["vbKeyZ".lower()] = 90
        
        self.globals["vbKey0".lower()] = 48
        self.globals["vbKey1".lower()] = 49
        self.globals["vbKey2".lower()] = 50
        self.globals["vbKey3".lower()] = 51
        self.globals["vbKey4".lower()] = 52
        self.globals["vbKey5".lower()] = 53
        self.globals["vbKey6".lower()] = 54
        self.globals["vbKey7".lower()] = 55
        self.globals["vbKey8".lower()] = 56
        self.globals["vbKey9".lower()] = 57
        
        self.globals["vbKeyNumpad0".lower()] = 0x60
        self.globals["vbKeyNumpad1".lower()] = 0x61
        self.globals["vbKeyNumpad2".lower()] = 0x62
        self.globals["vbKeyNumpad3".lower()] = 0x63
        self.globals["vbKeyNumpad4".lower()] = 0x64
        self.globals["vbKeyNumpad5".lower()] = 0x65
        self.globals["vbKeyNumpad6".lower()] = 0x66
        self.globals["vbKeyNumpad7".lower()] = 0x67
        self.globals["vbKeyNumpad8".lower()] = 0x68
        self.globals["vbKeyNumpad9".lower()] = 0x69
        self.globals["vbKeyMultiply".lower()] = 0x6A
        self.globals["vbKeyAdd".lower()] = 0x6B
        self.globals["vbKeySeparator".lower()] = 0x6C
        self.globals["vbKeySubtract".lower()] = 0x6D
        self.globals["vbKeyDecimal".lower()] = 0x6E
        self.globals["vbKeyDivide".lower()] = 0x6F
        
        self.globals["vbKeyF1".lower()] = 0x70
        self.globals["vbKeyF2".lower()] = 0x71
        self.globals["vbKeyF3".lower()] = 0x72
        self.globals["vbKeyF4".lower()] = 0x73
        self.globals["vbKeyF5".lower()] = 0x74
        self.globals["vbKeyF6".lower()] = 0x75
        self.globals["vbKeyF7".lower()] = 0x76
        self.globals["vbKeyF8".lower()] = 0x77
        self.globals["vbKeyF9".lower()] = 0x78
        self.globals["vbKeyF10".lower()] = 0x79
        self.globals["vbKeyF11".lower()] = 0x7A
        self.globals["vbKeyF12".lower()] = 0x7B
        self.globals["vbKeyF13".lower()] = 0x7C
        self.globals["vbKeyF14".lower()] = 0x7D
        self.globals["vbKeyF15".lower()] = 0x7E
        self.globals["vbKeyF16".lower()] = 0x7F        

        self.globals["VBA.vbKeyLButton".lower()] = 0x1
        self.globals["VBA.vbKeyRButton".lower()] = 0x2
        self.globals["VBA.vbKeyCancel".lower()] = 0x3
        self.globals["VBA.vbKeyMButton".lower()] = 0x4
        self.globals["VBA.vbKeyBack".lower()] = 0x8
        self.globals["VBA.vbKeyTab".lower()] = 0x9
        self.globals["VBA.vbKeyClear".lower()] = 0xC
        self.globals["VBA.vbKeyReturn".lower()] = 0xD
        self.globals["VBA.vbKeyShift".lower()] = 0x10
        self.globals["VBA.vbKeyControl".lower()] = 0x11
        self.globals["VBA.vbKeyMenu".lower()] = 0x12
        self.globals["VBA.vbKeyPause".lower()] = 0x13
        self.globals["VBA.vbKeyCapital".lower()] = 0x14
        self.globals["VBA.vbKeyEscape".lower()] = 0x1B
        self.globals["VBA.vbKeySpace".lower()] = 0x20
        self.globals["VBA.vbKeyPageUp".lower()] = 0x21
        self.globals["VBA.vbKeyPageDown".lower()] = 0x22
        self.globals["VBA.vbKeyEnd".lower()] = 0x23
        self.globals["VBA.vbKeyHome".lower()] = 0x24
        self.globals["VBA.vbKeyLeft".lower()] = 0x25
        self.globals["VBA.vbKeyUp".lower()] = 0x26
        self.globals["VBA.vbKeyRight".lower()] = 0x27
        self.globals["VBA.vbKeyDown".lower()] = 0x28
        self.globals["VBA.vbKeySelect".lower()] = 0x29
        self.globals["VBA.vbKeyPrint".lower()] = 0x2A
        self.globals["VBA.vbKeyExecute".lower()] = 0x2B
        self.globals["VBA.vbKeySnapshot".lower()] = 0x2C
        self.globals["VBA.vbKeyInsert".lower()] = 0x2D
        self.globals["VBA.vbKeyDelete".lower()] = 0x2E
        self.globals["VBA.vbKeyHelp".lower()] = 0x2F
        self.globals["VBA.vbKeyNumlock".lower()] = 0x90

        self.globals["VBA.vbKeyA".lower()] = 65
        self.globals["VBA.vbKeyB".lower()] = 66
        self.globals["VBA.vbKeyC".lower()] = 67
        self.globals["VBA.vbKeyD".lower()] = 68
        self.globals["VBA.vbKeyE".lower()] = 69
        self.globals["VBA.vbKeyF".lower()] = 70
        self.globals["VBA.vbKeyG".lower()] = 71
        self.globals["VBA.vbKeyH".lower()] = 72
        self.globals["VBA.vbKeyI".lower()] = 73
        self.globals["VBA.vbKeyJ".lower()] = 74
        self.globals["VBA.vbKeyK".lower()] = 75
        self.globals["VBA.vbKeyL".lower()] = 76
        self.globals["VBA.vbKeyM".lower()] = 77
        self.globals["VBA.vbKeyN".lower()] = 78
        self.globals["VBA.vbKeyO".lower()] = 79
        self.globals["VBA.vbKeyP".lower()] = 80
        self.globals["VBA.vbKeyQ".lower()] = 81
        self.globals["VBA.vbKeyR".lower()] = 82
        self.globals["VBA.vbKeyS".lower()] = 83
        self.globals["VBA.vbKeyT".lower()] = 84
        self.globals["VBA.vbKeyU".lower()] = 85
        self.globals["VBA.vbKeyV".lower()] = 86
        self.globals["VBA.vbKeyW".lower()] = 87
        self.globals["VBA.vbKeyX".lower()] = 88
        self.globals["VBA.vbKeyY".lower()] = 89
        self.globals["VBA.vbKeyZ".lower()] = 90
        
        self.globals["VBA.vbKey0".lower()] = 48
        self.globals["VBA.vbKey1".lower()] = 49
        self.globals["VBA.vbKey2".lower()] = 50
        self.globals["VBA.vbKey3".lower()] = 51
        self.globals["VBA.vbKey4".lower()] = 52
        self.globals["VBA.vbKey5".lower()] = 53
        self.globals["VBA.vbKey6".lower()] = 54
        self.globals["VBA.vbKey7".lower()] = 55
        self.globals["VBA.vbKey8".lower()] = 56
        self.globals["VBA.vbKey9".lower()] = 57
        
        self.globals["VBA.vbKeyNumpad0".lower()] = 0x60
        self.globals["VBA.vbKeyNumpad1".lower()] = 0x61
        self.globals["VBA.vbKeyNumpad2".lower()] = 0x62
        self.globals["VBA.vbKeyNumpad3".lower()] = 0x63
        self.globals["VBA.vbKeyNumpad4".lower()] = 0x64
        self.globals["VBA.vbKeyNumpad5".lower()] = 0x65
        self.globals["VBA.vbKeyNumpad6".lower()] = 0x66
        self.globals["VBA.vbKeyNumpad7".lower()] = 0x67
        self.globals["VBA.vbKeyNumpad8".lower()] = 0x68
        self.globals["VBA.vbKeyNumpad9".lower()] = 0x69
        self.globals["VBA.vbKeyMultiply".lower()] = 0x6A
        self.globals["VBA.vbKeyAdd".lower()] = 0x6B
        self.globals["VBA.vbKeySeparator".lower()] = 0x6C
        self.globals["VBA.vbKeySubtract".lower()] = 0x6D
        self.globals["VBA.vbKeyDecimal".lower()] = 0x6E
        self.globals["VBA.vbKeyDivide".lower()] = 0x6F
        
        self.globals["VBA.vbKeyF1".lower()] = 0x70
        self.globals["VBA.vbKeyF2".lower()] = 0x71
        self.globals["VBA.vbKeyF3".lower()] = 0x72
        self.globals["VBA.vbKeyF4".lower()] = 0x73
        self.globals["VBA.vbKeyF5".lower()] = 0x74
        self.globals["VBA.vbKeyF6".lower()] = 0x75
        self.globals["VBA.vbKeyF7".lower()] = 0x76
        self.globals["VBA.vbKeyF8".lower()] = 0x77
        self.globals["VBA.vbKeyF9".lower()] = 0x78
        self.globals["VBA.vbKeyF10".lower()] = 0x79
        self.globals["VBA.vbKeyF11".lower()] = 0x7A
        self.globals["VBA.vbKeyF12".lower()] = 0x7B
        self.globals["VBA.vbKeyF13".lower()] = 0x7C
        self.globals["VBA.vbKeyF14".lower()] = 0x7D
        self.globals["VBA.vbKeyF15".lower()] = 0x7E
        self.globals["VBA.vbKeyF16".lower()] = 0x7F        

        # Excel error codes.
        self.globals["xlErrDiv0".lower()] = 2007  #DIV/0!
        self.globals["xlErrNA".lower()] = 2042    #N/A
        self.globals["xlErrName".lower()] = 2029  #NAME?
        self.globals["xlErrNull".lower()] = 2000  #NULL!
        self.globals["xlErrNum".lower()] = 2036   #NUM!
        self.globals["xlErrRef".lower()] = 2023   #REF!
        self.globals["xlErrValue".lower()] = 2015 #VALUE!

        # System info.
        self.globals["System.OperatingSystem".lower()] = "Windows NT"

        # Call type constants.
        self.globals["vbGet".lower()] = 2
        self.globals["vbLet".lower()] = 4
        self.globals["vbMethod".lower()] = 1
        self.globals["vbSet".lower()] = 8

        # XlTickMark Enum
        self.globals["xlTickMarkCross".lower()] = 4	
        self.globals["xlTickMarkInside".lower()] = 2	
        self.globals["xlTickMarkNone".lower()] = -4142	
        self.globals["xlTickMarkOutside".lower()] = 3	

        # XlXmlExportResult Enum
        self.globals["xlXmlExportSuccess".lower()] = 0	
        self.globals["xlXmlExportValidationFailed".lower()] = 1	

        # Type
        self.globals["xlCellValue".lower()] = 1 
        self.globals["xlExpression".lower()] = 2
        self.globals["xlColorScale".lower()] = 3
        self.globals["xlDatabar".lower()] = 4
        self.globals["xlTop10".lower()] = 5
        self.globals["xlIconSet".lower()] = 6
        self.globals["xlUniqueValues".lower()] = 8
        self.globals["xlTextString".lower()] = 9
        self.globals["xlBlanksCondition".lower()] = 10
        self.globals["xlTimePeriod".lower()] = 11
        self.globals["xlAboveAverageCondition".lower()] = 12
        self.globals["xlNoBlanksCondition".lower()] = 13
        self.globals["xlErrorsCondition".lower()] = 16
        self.globals["xlNoErrorsCondition".lower()] = 17

        # Operator
        self.globals["xlBetween".lower()] = 1
        self.globals["xlNotBetween".lower()] = 2
        self.globals["xlEqual".lower()] = 3
        self.globals["xlNotEqual".lower()] = 4
        self.globals["xlGreater".lower()] = 5
        self.globals["xlLess".lower()] = 6
        self.globals["xlGreaterEqual".lower()] = 7
        self.globals["xlLessEqual".lower()] = 8

        # CartType
        self.globals["xl3DArea".lower()] = -4098
        self.globals["xl3DAreaStacked".lower()] = 78
        self.globals["xl3DAreaStacked100".lower()] = 79
        self.globals["xl3DBarClustered".lower()] = 60
        self.globals["xl3DBarStacked".lower()] = 61
        self.globals["xl3DBarStacked100".lower()] = 62
        self.globals["xl3DColumn".lower()] = -4100
        self.globals["xl3DColumnClustered".lower()] = 54
        self.globals["xl3DColumnStacked".lower()] = 55
        self.globals["xl3DColumnStacked100".lower()] = 56
        self.globals["xl3DLine".lower()] = -4101
        self.globals["xl3DPie".lower()] = -4102
        self.globals["xl3DPieExploded".lower()] = 70
        self.globals["xlArea".lower()] = 1
        self.globals["xlAreaStacked".lower()] = 76
        self.globals["xlAreaStacked100".lower()] = 77
        self.globals["xlBarClustered".lower()] = 57
        self.globals["xlBarOfPie".lower()] = 71
        self.globals["xlBarStacked".lower()] = 58
        self.globals["xlBarStacked100".lower()] = 59
        self.globals["xlBubble".lower()] = 15
        self.globals["xlBubble3DEffect".lower()] = 87
        self.globals["xlColumnClustered".lower()] = 51
        self.globals["xlColumnStacked".lower()] = 52
        self.globals["xlColumnStacked100".lower()] = 53
        self.globals["xlConeBarClustered".lower()] = 102
        self.globals["xlConeBarStacked".lower()] = 103
        self.globals["xlConeBarStacked100".lower()] = 104
        self.globals["xlConeCol".lower()] = 105
        self.globals["xlConeColClustered".lower()] = 99
        self.globals["xlConeColStacked".lower()] = 100
        self.globals["xlConeColStacked100".lower()] = 101
        self.globals["xlCylinderBarClustered".lower()] = 95
        self.globals["xlCylinderBarStacked".lower()] = 96
        self.globals["xlCylinderBarStacked100".lower()] = 97
        self.globals["xlCylinderCol".lower()] = 98
        self.globals["xlCylinderColClustered".lower()] = 92
        self.globals["xlCylinderColStacked".lower()] = 93
        self.globals["xlCylinderColStacked100".lower()] = 94
        self.globals["xlDoughnut".lower()] = -4120
        self.globals["xlDoughnutExploded".lower()] = 80
        self.globals["xlLine".lower()] = 4
        self.globals["xlLineMarkers".lower()] = 65
        self.globals["xlLineMarkersStacked".lower()] = 66
        self.globals["xlLineMarkersStacked100".lower()] = 67
        self.globals["xlLineStacked".lower()] = 63
        self.globals["xlLineStacked100".lower()] = 64
        self.globals["xlPie".lower()] = 5
        self.globals["xlPieExploded".lower()] = 69
        self.globals["xlPieOfPie".lower()] = 68
        self.globals["xlPyramidBarClustered".lower()] = 109
        self.globals["xlPyramidBarStacked".lower()] = 110
        self.globals["xlPyramidBarStacked100".lower()] = 111
        self.globals["xlPyramidCol".lower()] = 112
        self.globals["xlPyramidColClustered".lower()] = 106
        self.globals["xlPyramidColStacked".lower()] = 107
        self.globals["xlPyramidColStacked100".lower()] = 108
        self.globals["xlRadar".lower()] = -4151
        self.globals["xlRadarFilled".lower()] = 82
        self.globals["xlRadarMarkers".lower()] = 81
        self.globals["xlStockHLC".lower()] = 88
        self.globals["xlStockOHLC".lower()] = 89
        self.globals["xlStockVHLC".lower()] = 90
        self.globals["xlStockVOHLC".lower()] = 91
        self.globals["xlSurface".lower()] = 83
        self.globals["xlSurfaceTopView".lower()] = 85
        self.globals["xlSurfaceTopViewWireframe".lower()] = 86
        self.globals["xlSurfaceWireframe".lower()] = 84
        self.globals["xlXYScatter".lower()] = -4169
        self.globals["xlXYScatterLines".lower()] = 74
        self.globals["xlXYScatterLinesNoMarkers".lower()] = 75
        self.globals["xlXYScatterSmooth".lower()] = 72
        self.globals["xlXYScatterSmoothNoMarkers".lower()] = 73

        # consts
        self.globals["xl3DBar".lower()] = -4099
        self.globals["xl3DEffects1".lower()] = 13
        self.globals["xl3DEffects2".lower()] = 14
        self.globals["xl3DSurface".lower()] = -4103
        self.globals["xlAbove".lower()] = 0
        self.globals["xlAccounting1".lower()] = 4
        self.globals["xlAccounting2".lower()] = 5
        self.globals["xlAccounting4".lower()] = 17
        self.globals["xlAdd".lower()] = 2
        self.globals["xlAll".lower()] = -4104
        self.globals["xlAccounting3".lower()] = 6
        self.globals["xlAllExceptBorders".lower()] = 7
        self.globals["xlAutomatic".lower()] = -4105
        self.globals["xlBar".lower()] = 2
        self.globals["xlBelow".lower()] = 1
        self.globals["xlBidi".lower()] = -5000
        self.globals["xlBidiCalendar".lower()] = 3
        self.globals["xlBoth".lower()] = 1
        self.globals["xlBottom".lower()] = -4107
        self.globals["xlCascade".lower()] = 7
        self.globals["xlCenter".lower()] = -4108
        self.globals["xlCenterAcrossSelection".lower()] = 7
        self.globals["xlChart4".lower()] = 2
        self.globals["xlChartSeries".lower()] = 17
        self.globals["xlChartShort".lower()] = 6
        self.globals["xlChartTitles".lower()] = 18
        self.globals["xlChecker".lower()] = 9
        self.globals["xlCircle".lower()] = 8
        self.globals["xlClassic1".lower()] = 1
        self.globals["xlClassic2".lower()] = 2
        self.globals["xlClassic3".lower()] = 3
        self.globals["xlClosed".lower()] = 3
        self.globals["xlColor1".lower()] = 7
        self.globals["xlColor2".lower()] = 8
        self.globals["xlColor3".lower()] = 9
        self.globals["xlColumn".lower()] = 3
        self.globals["xlCombination".lower()] = -4111
        self.globals["xlComplete".lower()] = 4
        self.globals["xlConstants".lower()] = 2
        self.globals["xlContents".lower()] = 2
        self.globals["xlContext".lower()] = -5002
        self.globals["xlCorner".lower()] = 2
        self.globals["xlCrissCross".lower()] = 16
        self.globals["xlCross".lower()] = 4
        self.globals["xlCustom".lower()] = -4114
        self.globals["xlDebugCodePane".lower()] = 13
        self.globals["xlDefaultAutoFormat".lower()] = -1
        self.globals["xlDesktop".lower()] = 9
        self.globals["xlDiamond".lower()] = 2
        self.globals["xlDirect".lower()] = 1
        self.globals["xlDistributed".lower()] = -4117
        self.globals["xlDivide".lower()] = 5
        self.globals["xlDoubleAccounting".lower()] = 5
        self.globals["xlDoubleClosed".lower()] = 5
        self.globals["xlDoubleOpen".lower()] = 4
        self.globals["xlDoubleQuote".lower()] = 1
        self.globals["xlDrawingObject".lower()] = 14
        self.globals["xlEntireChart".lower()] = 20
        self.globals["xlExcelMenus".lower()] = 1
        self.globals["xlExtended".lower()] = 3
        self.globals["xlFill".lower()] = 5
        self.globals["xlFirst".lower()] = 0
        self.globals["xlFixedValue".lower()] = 1
        self.globals["xlFloating".lower()] = 5
        self.globals["xlFormats".lower()] = -4122
        self.globals["xlFormula".lower()] = 5
        self.globals["xlFullScript".lower()] = 1
        self.globals["xlGeneral".lower()] = 1
        self.globals["xlGray16".lower()] = 17
        self.globals["xlGray25".lower()] = -4124
        self.globals["xlGray50".lower()] = -4125
        self.globals["xlGray75".lower()] = -4126
        self.globals["xlGray8".lower()] = 18
        self.globals["xlGregorian".lower()] = 2
        self.globals["xlGrid".lower()] = 15
        self.globals["xlGridline".lower()] = 22
        self.globals["xlHigh".lower()] = -4127
        self.globals["xlHindiNumerals".lower()] = 3
        self.globals["xlIcons".lower()] = 1
        self.globals["xlImmediatePane".lower()] = 12
        self.globals["xlInside".lower()] = 2
        self.globals["xlInteger".lower()] = 2
        self.globals["xlJustify".lower()] = -4130
        self.globals["xlLast".lower()] = 1
        self.globals["xlLastCell".lower()] = 11
        self.globals["xlLatin".lower()] = -5001
        self.globals["xlLeft".lower()] = -4131
        self.globals["xlLeftToRight".lower()] = 2
        self.globals["xlLightDown".lower()] = 13
        self.globals["xlLightHorizontal".lower()] = 11
        self.globals["xlLightUp".lower()] = 14
        self.globals["xlLightVertical".lower()] = 12
        self.globals["xlList1".lower()] = 10
        self.globals["xlList2".lower()] = 11
        self.globals["xlList3".lower()] = 12
        self.globals["xlLocalFormat1".lower()] = 15
        self.globals["xlLocalFormat2".lower()] = 16
        self.globals["xlLogicalCursor".lower()] = 1
        self.globals["xlLong".lower()] = 3
        self.globals["xlLotusHelp".lower()] = 2
        self.globals["xlLow".lower()] = -4134
        self.globals["xlLTR".lower()] = -5003
        self.globals["xlMacrosheetCell".lower()] = 7
        self.globals["xlManual".lower()] = -4135
        self.globals["xlMaximum".lower()] = 2
        self.globals["xlMinimum".lower()] = 4
        self.globals["xlMinusValues".lower()] = 3
        self.globals["xlMixed".lower()] = 2
        self.globals["xlMixedAuthorizedScript".lower()] = 4
        self.globals["xlMixedScript".lower()] = 3
        self.globals["xlModule".lower()] = -4141
        self.globals["xlMultiply".lower()] = 4
        self.globals["xlNarrow".lower()] = 1
        self.globals["xlNextToAxis".lower()] = 4
        self.globals["xlNoDocuments".lower()] = 3
        self.globals["xlNone".lower()] = -4142
        self.globals["xlNotes".lower()] = -4144
        self.globals["xlOff".lower()] = -4146
        self.globals["xlOn".lower()] = 1
        self.globals["xlOpaque".lower()] = 3
        self.globals["xlOpen".lower()] = 2
        self.globals["xlOutside".lower()] = 3
        self.globals["xlPartial".lower()] = 3
        self.globals["xlPartialScript".lower()] = 2
        self.globals["xlPercent".lower()] = 2
        self.globals["xlPlus".lower()] = 9
        self.globals["xlPlusValues".lower()] = 2
        self.globals["xlReference".lower()] = 4
        self.globals["xlRight".lower()] = -4152
        self.globals["xlRTL".lower()] = -5004
        self.globals["xlScale".lower()] = 3
        self.globals["xlSemiautomatic".lower()] = 2
        self.globals["xlSemiGray75".lower()] = 10
        self.globals["xlShort".lower()] = 1
        self.globals["xlShowLabel".lower()] = 4
        self.globals["xlShowLabelAndPercent".lower()] = 5
        self.globals["xlShowPercent".lower()] = 3
        self.globals["xlShowValue".lower()] = 2
        self.globals["xlSimple".lower()] = -4154
        self.globals["xlSingle".lower()] = 2
        self.globals["xlSingleAccounting".lower()] = 4
        self.globals["xlSingleQuote".lower()] = 2
        self.globals["xlSolid".lower()] = 1
        self.globals["xlSquare".lower()] = 1
        self.globals["xlStar".lower()] = 5
        self.globals["xlStError".lower()] = 4
        self.globals["xlStrict".lower()] = 2
        self.globals["xlSubtract".lower()] = 3
        self.globals["xlSystem".lower()] = 1
        self.globals["xlTextBox".lower()] = 16
        self.globals["xlTiled".lower()] = 1
        self.globals["xlTitleBar".lower()] = 8
        self.globals["xlToolbar".lower()] = 1
        self.globals["xlToolbarButton".lower()] = 2
        self.globals["xlTop".lower()] = -4160
        self.globals["xlTopToBottom".lower()] = 1
        self.globals["xlTransparent".lower()] = 2
        self.globals["xlTriangle".lower()] = 3
        self.globals["xlVeryHidden".lower()] = 2
        self.globals["xlVisible".lower()] = 12
        self.globals["xlVisualCursor".lower()] = 2
        self.globals["xlWatchPane".lower()] = 11
        self.globals["xlWide".lower()] = 3
        self.globals["xlWorkbookTab".lower()] = 6
        self.globals["xlWorksheet4".lower()] = 1
        self.globals["xlWorksheetCell".lower()] = 3
        self.globals["xlWorksheetShort".lower()] = 5

        # XlBinsType
        self.globals["xlBinsTypeAutomatic".lower()] = 0
        self.globals["xlBinsTypeCategorical".lower()] = 1
        self.globals["xlBinsTypeManual".lower()] = 2
        self.globals["xlBinsTypeBinSize".lower()] = 3
        self.globals["xlBinsTypeBinCount".lower()] = 4

        # Application.International
        self.globals["xlCountryCode".lower()] = 1
        self.globals["xlCountrySetting".lower()] = 2

        # Excel consts from MSDN
        self.globals["xl3DBar".lower()] = -4099
        self.globals["xlAboveAverage".lower()] = 0
        self.globals["xlActionTypeDrillthrough".lower()] = 256
        self.globals["xlAutomaticAllocation".lower()] = 2
        self.globals["xlEqualAllocation".lower()] = 1
        self.globals["xlAllocateIncrement".lower()] = 2
        self.globals["xl24HourClock".lower()] = 33
        self.globals["xlColumnThenRow".lower()] = 2
        self.globals["xlArabicBothStrict".lower()] = 3
        self.globals["xlArrangeStyleCascade".lower()] = 7
        self.globals["xlArrowHeadLengthLong".lower()] = 3
        self.globals["xlArrowHeadStyleClosed".lower()] = 3
        self.globals["xlArrowHeadWidthMedium".lower()] = -4138
        self.globals["xlFillCopy".lower()] = 1
        self.globals["xlAnd".lower()] = 1
        self.globals["xlAxisCrossesAutomatic".lower()] = -4105
        self.globals["xlPrimary".lower()] = 1
        self.globals["xlCategory".lower()] = 1
        self.globals["xlBackgroundAutomatic".lower()] = -4105
        self.globals["xlBox".lower()] = 0
        self.globals["xlBinsTypeAutomatic".lower()] = 0
        self.globals["xlHairline".lower()] = 1
        self.globals["xlDiagonalDown".lower()] = 5
        self.globals["xlDialogActivate".lower()] = 103
        self.globals["xlErrDiv0".lower()] = 2007
        self.globals["xlAllValues".lower()] = 0
        self.globals["xlCalculatedMeasure".lower()] = 2
        self.globals["xlCalculationAutomatic".lower()] = -4105
        self.globals["xlAnyKey".lower()] = 2
        self.globals["xlCalculating".lower()] = 1
        self.globals["xlAutomaticScale".lower()] = -4105
        self.globals["xlCellChangeApplied".lower()] = 3
        self.globals["xlInsertDeleteCells".lower()] = 1
        self.globals["xlCellTypeAllFormatConditions".lower()] = -4172
        self.globals["xlChartElementPositionAutomatic".lower()] = -4105
        self.globals["xlAnyGallery".lower()] = 23
        self.globals["xlAxis".lower()] = 21
        self.globals["xlLocationAsNewSheet".lower()] = 1
        self.globals["xlAllFaces".lower()] = 7
        self.globals["xlStack".lower()] = 2
        self.globals["xlSplitByCustomSplit".lower()] = 4
        self.globals["xl3DArea".lower()] = -4098
        self.globals["xlCheckInMajorVersion".lower()] = 1
        self.globals["xlClipboardFormatBIFF".lower()] = 8
        self.globals["xlCmdCube".lower()] = 1
        self.globals["xlColorIndexAutomatic".lower()] = -4105
        self.globals["xlDMYFormat".lower()] = 4
        self.globals["xlCommandUnderlinesAutomatic".lower()] = -4105
        self.globals["xlCommentAndIndicator".lower()] = 1
        self.globals["xlConditionValueAutomaticMax".lower()] = 7
        self.globals["xlConnectionTypeDATAFEED".lower()] = 6
        self.globals["xlAverage".lower()] = -4106
        self.globals["xlBeginsWith".lower()] = 2
        self.globals["xlBitmap".lower()] = 2
        self.globals["xlExtractData".lower()] = 2
        self.globals["xlCreatorCode".lower()] = 1480803660
        self.globals["CredentialsMethodIntegrated".lower()] = 0
        self.globals["xlCubeAttribute".lower()] = 4
        self.globals["xlHierarchy".lower()] = 1
        self.globals["xlCopy".lower()] = 1
        self.globals["xlValidAlertInformation".lower()] = 3
        self.globals["xlValidateCustom".lower()] = 7
        self.globals["xlDataBarAxisAutomatic".lower()] = 0
        self.globals["xlDataBarBorderNone".lower()] = 0
        self.globals["xlDataBarFillGradient".lower()] = 1
        self.globals["xlDataBarColor".lower()] = 0
        self.globals["xlLabelPositionAbove".lower()] = 0
        self.globals["xlDataLabelSeparatorDefault".lower()] = 1
        self.globals["xlDataLabelsShowBubbleSizes".lower()] = 6
        self.globals["xlDay".lower()] = 1
        self.globals["xlAutoFill".lower()] = 4
        self.globals["xlShiftToLeft".lower()] = -4159
        self.globals["xlDown".lower()] = -4121
        self.globals["xlInterpolated".lower()] = 3
        self.globals["xlDisplayShapes".lower()] = -4104
        self.globals["xlHundredMillions".lower()] = -8
        self.globals["xlDuplicate".lower()] = 1
        self.globals["xlFilterAboveAverage".lower()] = 33
        self.globals["xlBIFF".lower()] = 2
        self.globals["xlAutomaticUpdate".lower()] = 4
        self.globals["xlPublisher".lower()] = 1
        self.globals["xlDisabled".lower()] = 0
        self.globals["xlNoRestrictions".lower()] = 0
        self.globals["xlCap".lower()] = 1
        self.globals["xlX".lower()] = -4168
        self.globals["xlErrorBarIncludeBoth".lower()] = 1
        self.globals["xlErrorBarTypeCustom".lower()] = -4114
        self.globals["xlEmptyCellReferences".lower()] = 7
        self.globals["xlReadOnly".lower()] = 3
        self.globals["xlAddIn".lower()] = 18
        self.globals["xlFileValidationPivotDefault".lower()] = 0
        self.globals["xlFillWithAll".lower()] = -4104
        self.globals["xlFilterCopy".lower()] = 2
        self.globals["xlFilterAllDatesInPeriodDay".lower()] = 2
        self.globals["xlFilterStatusOK".lower()] = 0
        self.globals["xlComments".lower()] = -4144
        self.globals["xlQualityMinimum".lower()] = 1
        self.globals["xlTypePDF".lower()] = 0
        self.globals["xlButtonControl".lower()] = 0
        self.globals["xlBetween".lower()] = 1
        self.globals["xlAboveAverageCondition".lower()] = 12
        self.globals["FilterBottom".lower()] = 0
        self.globals["xlColumnLabels".lower()] = 2
        self.globals["xlA1TableRefs".lower()] = 0
        self.globals["GradientFillLinear".lower()] = 0
        self.globals["xlHAlignCenter".lower()] = -4108
        self.globals["xlHebrewFullScript".lower()] = 0
        self.globals["xlAllChanges".lower()] = 2
        self.globals["xlHtmlCalc".lower()] = 1
        self.globals["xlIMEModeAlpha".lower()] = 8
        self.globals["xlIcon0Bars".lower()] = 37
        self.globals["xl3Arrows".lower()] = 1
        self.globals["xlPivotTableReport".lower()] = 1
        self.globals["xlFormatFromLeftOrAbove".lower()] = 0
        self.globals["xlShiftDown".lower()] = -4121
        self.globals["xlOutline".lower()] = 1
        self.globals["xlCompactRow".lower()] = 0
        self.globals["xlLegendPositionBottom".lower()] = -4107
        self.globals["xlContinuous".lower()] = 1
        self.globals["xlExcelLinks".lower()] = 1
        self.globals["xlEditionDate".lower()] = 2
        self.globals["xlLinkInfoOLELinks".lower()] = 2
        self.globals["xlLinkStatusCopiedValues".lower()] = 10
        self.globals["xlLinkTypeExcelLinks".lower()] = 1
        self.globals["xlListConflictDialog".lower()] = 0
        self.globals["xlListDataTypeCheckbox".lower()] = 9
        self.globals["xlSrcExternal".lower()] = 0
        self.globals["xlColumnHeader".lower()] = -4110
        self.globals["xlPart".lower()] = 2
        self.globals["LookForBlanks".lower()] = 0
        self.globals["xlMicrosoftAccess".lower()] = 4
        self.globals["xlMAPI".lower()] = 1
        self.globals["xlMarkerStyleAutomatic".lower()] = -4105
        self.globals["xlCentimeters".lower()] = 1
        self.globals["xlChangeByExcel".lower()] = 0
        self.globals["xlNoButton".lower()] = 0
        self.globals["xlDefault".lower()] = -4143
        self.globals["xlOLEControl".lower()] = 2
        self.globals["xlVerbOpen".lower()] = 2
        self.globals["xlOartHorizontalOverflowClip".lower()] = 1
        self.globals["xlOartVerticalOverflowClip".lower()] = 1
        self.globals["xlFitToPage".lower()] = 2
        self.globals["xlDownThenOver".lower()] = 1
        self.globals["xlDownward".lower()] = -4170
        self.globals["xlBlanks".lower()] = 4
        self.globals["xlPageBreakAutomatic".lower()] = -4105
        self.globals["xlPageBreakFull".lower()] = 1
        self.globals["xlLandscape".lower()] = 2
        self.globals["xlPaper10x14".lower()] = 16
        self.globals["xlParamTypeBigInt".lower()] = -5
        self.globals["xlConstant".lower()] = 1
        self.globals["xlPasteSpecialOperationAdd".lower()] = 2
        self.globals["xlPasteAll".lower()] = -4104
        self.globals["xlPatternAutomatic".lower()] = -4105
        self.globals["xlPhoneticAlignCenter".lower()] = 2
        self.globals["xlHiragana".lower()] = 2
        self.globals["xlPrinter".lower()] = 2
        self.globals["xlBMP".lower()] = 1
        self.globals["xlCenterPoint".lower()] = 5
        self.globals["xlHorizontalCoordinate".lower()] = 1
        self.globals["xlPivotCellBlankCell".lower()] = 9
        self.globals["xlDataFieldScope".lower()] = 2
        self.globals["xlDifferenceFrom".lower()] = 2
        self.globals["xlDate".lower()] = 2
        self.globals["xlColumnField".lower()] = 2
        self.globals["xlDoNotRepeatLabels".lower()] = 1
        self.globals["xlBefore".lower()] = 31
        self.globals["xlPTClassic".lower()] = 20
        self.globals["xlPivotLineBlank".lower()] = 3
        self.globals["xlMissingItemsDefault".lower()] = -1
        self.globals["xlConsolidation".lower()] = 3
        self.globals["xlPivotTableVersion2000".lower()] = 0
        self.globals["xlFreeFloating".lower()] = 3
        self.globals["xlMacintosh".lower()] = 1
        self.globals["xlPortugueseBoth".lower()] = 3
        self.globals["xlPrintErrorsBlank".lower()] = 1
        self.globals["xlPrintInPlace".lower()] = 16
        self.globals["xlPriorityHigh".lower()] = -4127
        self.globals["xlDisplayPropertyInPivotTable".lower()] = 1
        self.globals["xlProtectedViewCloseEdit".lower()] = 1
        self.globals["xlProtectedViewWindowMaximized".lower()] = 2
        self.globals["xlADORecordset".lower()] = 7
        self.globals["xlLensOnly".lower()] = 0
        self.globals["xlRangeAutoFormat3DEffects1".lower()] = 13
        self.globals["xlRangeValueDefault".lower()] = 10
        self.globals["xlA1".lower()] = 1
        self.globals["xlAbsolute".lower()] = 1
        self.globals["xlRDIAll".lower()] = 99
        self.globals["rgbAliceBlue".lower()] = 16775408
        self.globals["xlAlways".lower()] = 1
        self.globals["xlColumns".lower()] = 2
        self.globals["xlAutoActivate".lower()] = 3
        self.globals["xlDoNotSaveChanges".lower()] = 2
        self.globals["xlExclusive".lower()] = 3
        self.globals["xlLocalSessionChanges".lower()] = 2
        self.globals["xlScaleLinear".lower()] = -4132
        self.globals["xlNext".lower()] = 1
        self.globals["xlByColumns".lower()] = 2
        self.globals["xlWithinSheet".lower()] = 1
        self.globals["xlChart".lower()] = -4109
        self.globals["xlSheetHidden".lower()] = 0
        self.globals["xlSizeIsArea".lower()] = 1
        self.globals["xlSlicer".lower()] = 1
        self.globals["xlSlicerCrossFilterHideButtonsWithNoData".lower()] = 4
        self.globals["xlSlicerSortAscending".lower()] = 2
        self.globals["xlSortNormal".lower()] = 0
        self.globals["xlPinYin".lower()] = 1
        self.globals["xlCodePage".lower()] = 2
        self.globals["SortOnCellColor".lower()] = 1
        self.globals["xlAscending".lower()] = 1
        self.globals["xlSortColumns".lower()] = 1
        self.globals["xlSortLabels".lower()] = 2
        self.globals["xlSourceAutoFilter".lower()] = 3
        self.globals["xlSpanishTuteoAndVoseo".lower()] = 1
        self.globals["xlSparkScaleCustom".lower()] = 3
        self.globals["xlSparkColumn".lower()] = 2
        self.globals["SparklineColumnsSquare".lower()] = 2
        self.globals["xlSpeakByColumns".lower()] = 1
        self.globals["xlErrors".lower()] = 16
        self.globals["ColorScaleBlackWhite".lower()] = 3
        self.globals["xlSubscribeToPicture".lower()] = -4147
        self.globals["xlAtBottom".lower()] = 2
        self.globals["xlSummaryOnLeft".lower()] = -4131
        self.globals["xlStandardSummary".lower()] = 1
        self.globals["xlSummaryAbove".lower()] = 0
        self.globals["xlTabPositionFirst".lower()] = 0
        self.globals["xlBlankRow".lower()] = 19
        self.globals["xlDelimited".lower()] = 1
        self.globals["xlTextQualifierDoubleQuote".lower()] = 1
        self.globals["xlTextVisualLTR".lower()] = 1
        self.globals["xlThemeColorAccent1".lower()] = 5
        self.globals["xlThemeFontMajor".lower()] = 2
        self.globals["xlThreadModeAutomatic".lower()] = 0
        self.globals["xlTickLabelOrientationAutomatic".lower()] = -4105
        self.globals["xlTickLabelPositionHigh".lower()] = -4127
        self.globals["xlTickMarkCross".lower()] = 4
        self.globals["xlLast7Days".lower()] = 2
        self.globals["xlDays".lower()] = 0
        self.globals["xlTimelineLevelYears".lower()] = 0
        self.globals["xlNoButtonChanges".lower()] = 1
        self.globals["xlTop10Bottom".lower()] = 0
        self.globals["xlTotalsCalculationAverage".lower()] = 2
        self.globals["xlExponential".lower()] = 5
        self.globals["xlUnderlineStyleDouble".lower()] = -4119
        self.globals["xlUpdateLinksAlways".lower()] = 3
        self.globals["xlVAlignBottom".lower()] = -4107
        self.globals["xlWBATChart".lower()] = -4109
        self.globals["xlWebFormattingAll".lower()] = 1
        self.globals["xlAllTables".lower()] = 2
        self.globals["xlMaximized".lower()] = -4137
        self.globals["xlChartAsWindow".lower()] = 5
        self.globals["xlNormalView".lower()] = 1
        self.globals["xlCommand".lower()] = 2
        self.globals["xlXmlExportSuccess".lower()] = 0
        self.globals["xlXmlImportElementsTruncated".lower()] = 1
        self.globals["xlXmlLoadImportToList".lower()] = 2
        self.globals["xlGuess".lower()] = 0
        self.globals["xlNumberFormatTypeDefault".lower()] = 0
        self.globals["xlNumberFormatTypeNumber".lower()] = 1
        self.globals["xlNumberFormatTypePercent".lower()] = 2
        self.globals["xlCategoryLabelLevelAll".lower()] = -1
        self.globals["xlCategoryLabelLevelCustom".lower()] = -2
        self.globals["xlCategoryLabelLevelNone".lower()] = -3
        self.globals["xlCategoryLabelLevelAll".lower()] = -1
        self.globals["xlCategoryLabelLevelCustom".lower()] = -2
        self.globals["xlCategoryLabelLevelNone".lower()] = -3
        self.globals["xlChartElementPositionAutomatic".lower()] = -4105
        self.globals["xlChartElementPositionCustom".lower()] = -4114
        self.globals["xlColorIndexAutomatic".lower()] = -4105
        self.globals["xlColorIndexNone".lower()] = -4142
        self.globals["xlShiftToLeft".lower()] = -4159
        self.globals["xlShiftUp".lower()] = -4162
        self.globals["xlDown".lower()] = -4121
        self.globals["xlToLeft".lower()] = -4159
        self.globals["xlToRight".lower()] = -4161
        self.globals["xlUp".lower()] = -4162
        self.globals["xlForecastAggregationAverage".lower()] = 1
        self.globals["xlForecastAggregationCount".lower()] = 2
        self.globals["xlForecastAggregationCountA".lower()] = 3
        self.globals["xlForecastAggregationMax".lower()] = 4
        self.globals["xlForecastAggregationMedian".lower()] = 5
        self.globals["xlForecastAggregationMin".lower()] = 6
        self.globals["xlForecastAggregationSum".lower()] = 7
        self.globals["xlForecastDataCompletionInterpolate".lower()] = 1
        self.globals["xlForecastDataCompletionZeros".lower()] = 0
        self.globals["xlParentDataLabelOptionsBanner".lower()] = 1
        self.globals["xlParentDataLabelOptionsNone".lower()] = 0
        self.globals["xlParentDataLabelOptionsOverlapping".lower()] = 2
        self.globals["xlSeriesNameLevelAll".lower()] = -1
        self.globals["xlSeriesNameLevelCustom".lower()] = -2
        self.globals["xlSeriesNameLevelNone".lower()] = -3

        # PowerPoint consts from MSDN
        self.globals["msoAnimAccumulateAlways".lower()] = 2
        self.globals["msoAnimAdditiveAddBase".lower()] = 1
        self.globals["msoAnimAfterEffectDim".lower()] = 1
        self.globals["msoAnimateChartAllAtOnce".lower()] = 7
        self.globals["msoAnimCommandTypeCall".lower()] = 1
        self.globals["msoAnimDirectionAcross".lower()] = 18
        self.globals["msoAnimEffectAppear".lower()] = 1
        self.globals["msoAnimEffectAfterFreeze".lower()] = 1
        self.globals["msoAnimEffectRestartAlways".lower()] = 1
        self.globals["msoAnimFilterEffectSubtypeAcross".lower()] = 9
        self.globals["msoAnimFilterEffectTypeBarn".lower()] = 1
        self.globals["msoAnimColor".lower()] = 7
        self.globals["msoAnimTextUnitEffectByCharacter".lower()] = 1
        self.globals["msoAnimTriggerAfterPrevious".lower()] = 3
        self.globals["msoAnimTypeColor".lower()] = 2
        self.globals["msoClickStateAfterAllAnimations".lower()] = -2
        self.globals["ppActionEndShow".lower()] = 6
        self.globals["ppAdvanceModeMixed".lower()] = -2
        self.globals["ppAfterEffectDim".lower()] = 2
        self.globals["ppAlertsAll".lower()] = 2
        self.globals["ppArrangeCascade".lower()] = 2
        self.globals["ppAutoSizeMixed".lower()] = -2
        self.globals["ppBaselineAlignBaseline".lower()] = 1
        self.globals["ppBorderBottom".lower()] = 3
        self.globals["ppBulletMixed".lower()] = -2
        self.globals["ppCaseLower".lower()] = 2
        self.globals["ppAnimateByCategory".lower()] = 2
        self.globals["ppCheckInMajorVersion".lower()] = 1
        self.globals["ppAccent1".lower()] = 6
        self.globals["ppDateTimeddddMMMMddyyyy".lower()] = 2
        self.globals["ppDirectionLeftToRight".lower()] = 1
        self.globals["ppEffectAppear".lower()] = 3844
        self.globals["ppFarEastLineBreakLevelCustom".lower()] = 3
        self.globals["ppFixedFormatIntentPrint".lower()] = 2
        self.globals["ppFixedFormatTypePDF".lower()] = 2
        self.globals["ppFollowColorsMixed".lower()] = -2
        self.globals["ppFrameColorsBlackTextOnWhite".lower()] = 5
        self.globals["ppHorizontalGuide".lower()] = 1
        self.globals["ppHTMLAutodetect".lower()] = 4
        self.globals["ppIndentControlMixed".lower()] = -2
        self.globals["ppMediaTaskStatusNone".lower()] = 0
        self.globals["ppMediaTypeMixed".lower()] = -2
        self.globals["ppMouseClick".lower()] = 1
        self.globals["ppBulletAlphaLCParenBoth".lower()] = 8
        self.globals["ppAlignCenter".lower()] = 2
        self.globals["ppPasteBitmap".lower()] = 1
        self.globals["ppPlaceholderBitmap".lower()] = 9
        self.globals["ppPlaying".lower()] = 0
        self.globals["ppPrintBlackAndWhite".lower()] = 2
        self.globals["ppPrintHandoutHorizontalFirst".lower()] = 2
        self.globals["ppPrintOutputBuildSlides".lower()] = 7
        self.globals["ppPrintAll".lower()] = 1
        self.globals["ppProtectedViewCloseNormal".lower()] = 0
        self.globals["ppPublishAll".lower()] = 1
        self.globals["ppRDIAll".lower()] = 99
        self.globals["ppResampleMediaProfileCustom".lower()] = 1
        self.globals["ppRevisionInfoBaseline".lower()] = 1
        self.globals["ppSaveAsAddIn".lower()] = 8
        self.globals["ppSelectionNone".lower()] = 0
        self.globals["ppLayoutBlank".lower()] = 12
        self.globals["ppSlideShowManualAdvance".lower()] = 1
        self.globals["ppSlideShowPointerAlwaysHidden".lower()] = 3
        self.globals["ppShowAll".lower()] = 1
        self.globals["ppSlideShowBlackScreen".lower()] = 3
        self.globals["ppShowTypeKiosk".lower()] = 3
        self.globals["ppSlideSize35MM".lower()] = 4
        self.globals["ppSoundEffectsMixed".lower()] = -2
        self.globals["ppSoundFormatCDAudio".lower()] = 3
        self.globals["ppTabStopCenter".lower()] = 2
        self.globals["ppAnimateByAllLevels".lower()] = 16
        self.globals["ppBodyStyle".lower()] = 3
        self.globals["ppAnimateByCharacter".lower()] = 2
        self.globals["ppTransitionSpeedFast".lower()] = 3
        self.globals["ppUpdateOptionAutomatic".lower()] = 2
        self.globals["ppViewHandoutMaster".lower()] = 4
        self.globals["ppWindowMaximized".lower()] = 3
        self.globals["xlAxisCrossesAutomatic".lower()] = -4105
        self.globals["xlPrimary".lower()] = 1
        self.globals["xlCategory".lower()] = 1
        self.globals["xlBackgroundAutomatic".lower()] = -4105
        self.globals["xlBox".lower()] = 0
        self.globals["xlBinsTypeAutomatic".lower()] = 0
        self.globals["xlHairline".lower()] = 1
        self.globals["xlCategoryLabelLevelAll".lower()] = -1
        self.globals["xlAutomaticScale".lower()] = -4105
        self.globals["xlChartElementPositionAutomatic".lower()] = -4105
        self.globals["xlAnyGallery".lower()] = 23
        self.globals["xlAxis".lower()] = 21
        self.globals["xlAllFaces".lower()] = 7
        self.globals["xlStack".lower()] = 2
        self.globals["xlSplitByCustomSplit".lower()] = 4
        self.globals["xlColorIndexAutomatic".lower()] = -4105
        self.globals["xl3DBar".lower()] = -4099
        self.globals["xlBitmap".lower()] = 2
        self.globals["xlLabelPositionAbove".lower()] = 0
        self.globals["xlDataLabelSeparatorDefault".lower()] = 1
        self.globals["xlDataLabelsShowBubbleSizes".lower()] = 6
        self.globals["xlInterpolated".lower()] = 3
        self.globals["xlHundredMillions".lower()] = -8
        self.globals["xlCap".lower()] = 1
        self.globals["xlChartX".lower()] = -4168
        self.globals["xlErrorBarIncludeBoth".lower()] = 1
        self.globals["xlErrorBarTypeCustom".lower()] = -4114
        self.globals["xlHAlignCenter".lower()] = -4108
        self.globals["xlLegendPositionBottom".lower()] = -4107
        self.globals["xlContinuous".lower()] = 1
        self.globals["xlMarkerStyleAutomatic".lower()] = -4105
        self.globals["xlDownward".lower()] = -4170
        self.globals["xlParentDataLabelOptionsNone".lower()] = 0
        self.globals["xlPatternAutomatic".lower()] = -4105
        self.globals["xlPrinter".lower()] = 2
        self.globals["xlCenterPoint".lower()] = 5
        self.globals["xlCenterPoint".lower()] = 5
        self.globals["xlColumnField".lower()] = 2
        self.globals["xlContext".lower()] = -5002
        self.globals["xlAliceBlue".lower()] = 16775408
        self.globals["xlColumns".lower()] = 2
        self.globals["xlScaleLinear".lower()] = -4132
        self.globals["xlSeriesNameLevelAll".lower()] = -1
        self.globals["xlSizeIsArea".lower()] = 1
        self.globals["xlTickLabelOrientationAutomatic".lower()] = -4105
        self.globals["xlTickLabelPositionHigh".lower()] = -4127
        self.globals["xlTickMarkCross".lower()] = 4
        self.globals["xlDays".lower()] = 0
        self.globals["xlExponential".lower()] = 5
        self.globals["xlUnderlineStyleDouble".lower()] = -4119
        self.globals["xlVAlignBottom".lower()] = -4107


    def get_true_name(self, name):
        """
        Get the true name of an aliased function imported from a DLL.
        """
        if (name in self.dll_func_true_names):
            return self.dll_func_true_names[name]
        return None
        
    def open_file(self, fname):
        """
        Simulate opening a file.

        fname - The name of the file.
        """

        # Save that the file is opened.
        self.open_files[fname] = {}
        self.open_files[fname]["name"] = fname
        self.open_files[fname]["contents"] = []

    def dump_all_files(self):
        for fname in list(self.open_files.keys()):
            self.dump_file(fname)
        
    def dump_file(self, file_id):
        """
        Save the contents of a file dumped by the VBA to disk.

        file_id - The name of the file.
        """

        # Make sure the "file" exists.
        file_id = str(file_id)
        if (file_id not in self.open_files):
            log.error("File " + file_id + " not open. Cannot save.")
            return
        
        # Get the name of the file being closed.
        name = self.open_files[file_id]["name"].replace("#", "")
        log.info("Closing file " + name)
        
        # Get the data written to the file and track it.
        data = self.open_files[file_id]["contents"]
        self.closed_files[name] = data

        # Clear the file out of the open files.
        del self.open_files[file_id]

        # Save the hash of the written file.
        raw_data = array.array('B', data).tostring()
        h = sha256()
        h.update(raw_data)
        file_hash = h.hexdigest()
        self.report_action("Dropped File Hash", file_hash, 'File Name: ' + name)

        # TODO: Set a flag to control whether to dump file contents.

        # Dump out the file.
        if (out_dir is not None):

            # Make the dropped file directory if needed.
            if (not os.path.isdir(out_dir)):
                os.makedirs(out_dir)

            # Dump the file.
            try:

                # Get a unique name for the file.
                short_name = name
                start = 0
                if ('\\' in short_name):
                    start = short_name.rindex('\\') + 1
                if ('/' in short_name):
                    start = short_name.rindex('/') + 1
                short_name = out_dir + short_name[start:].strip()
                try:
                    f = open(short_name, 'r')
                    # Already exists. Get a unique name.
                    f.close()
                    file_count += 1
                    short_name += " (" + str(file_count) + ")"
                except:
                    pass
                    
                # Write out the dropped file.
                f = open(short_name, 'wb')
                f.write(raw_data)
                f.close()
                log.info("Wrote dumped file (hash " + file_hash + ") to " + short_name + " .")
                
            except Exception as e:
                log.error("Writing file " + short_name + " failed. " + str(e))

        else:
            log.warning("File not dumped. Output dir is None.")
        
    def _get(self, name):

        if (not isinstance(name, str)):
            raise KeyError('Object %r not found' % name)

        # convert to lowercase
        name = name.lower()
        log.debug("Looking for var '" + name + "'...")
        
        # First, search in locals. This handles variables whose name overrides
        # a system function.
        if name in self.locals:
            log.debug('Found %r in locals' % name)
            return self.locals[name]
        # second, in globals:
        elif name in self.globals:
            log.debug('Found %r in globals' % name)
            return self.globals[name]
        # next, search in the global VBA library:
        elif name in VBA_LIBRARY:
            log.debug('Found %r in VBA Library' % name)
            return VBA_LIBRARY[name]
        # Unknown symbol.
        else:            
            raise KeyError('Object %r not found' % name)
            # NOTE: if name is unknown, just raise Python dict's exception
            # TODO: raise a custom VBA exception?

    def get(self, name):

        # First try to get the item using the current with context.
        tmp_name = str(self.with_prefix) + str(name)
        try:
            return self._get(tmp_name)
        except KeyError:
            pass

        # Now try it without the current with context.
        try:
            return self._get(str(name))
        except KeyError:
            pass

        # Finally see if the variable was initially defined with a trailing '$'.
        return self._get(str(name) + "$")

    def contains(self, name, local=False):
        if (local):
            return (str(name).lower() in self.locals)
        try:
            self.get(name)
            return True
        except KeyError:
            return False

    def contains_user_defined(self, name):
        return ((name in self.locals) or (name in self.globals))
        
    def get_type(self, var):
        if (not isinstance(var, str)):
            return None
        var = var.lower()
        if (var not in self.types):
            return None
        return self.types[var]

    def get_doc_var(self, var):
        if (not isinstance(var, str)):
            return None
        var = var.lower()
        log.info("Looking up doc var " + var)
        if (var not in self.doc_vars):

            # Can't find a doc var with this name. See if we have an internal variable
            # with this name.
            log.debug("doc var named " + var + " not found.")
            try:
                var_value = self.get(var)
                if (var_value is not None):
                    return self.get_doc_var(var_value)
            except KeyError:
                pass

            # Can't find it. Do we have a wild card doc var to guess for
            # this value?
            if ("*" in self.doc_vars):
                return self.doc_vars["*"]

            # No wildcard variable. Return nothing.
            return None

        # Found it.
        r = self.doc_vars[var]
        log.debug("Found doc var " + var + " = " + str(r))
        return r
            
    # TODO: set_global?

    def set(self, name, value, var_type=None, do_with_prefix=True):
        if (not isinstance(name, str)):
            return
        # convert to lowercase
        name = name.lower()
        if name in self.locals:
            self.locals[name] = value
        # check globals, but avoid to overwrite subs and functions:
        elif name in self.globals and not is_procedure(self.globals[name]):
            self.globals[name] = value
            log.debug("Set global var " + name + " = " + str(value))
        else:
            # new name, always stored in locals:
            self.locals[name] = value

        # If we know the type of the variable, save it.
        if (var_type is not None):
            self.types[name] = var_type

        # Also set the variable using the current With name prefix, if
        # we have one.
        if ((do_with_prefix) and (len(self.with_prefix) > 0)):
            tmp_name = str(self.with_prefix) + str(name)
            self.set(tmp_name, value, var_type=var_type, do_with_prefix=False)

        # Handle base64 conversion with VBA objects.
        if (name.endswith(".text")):

            # Is the root object something set to the "bin.base64" data type?
            node_type = name.replace(".text", ".datatype")
            try:
                val = str(self.get(node_type)).strip()
                if (val == "bin.base64"):

                    # Try converting the text from base64.
                    try:

                        # Set the typed vale of the node to the decoded value.
                        conv_val = base64.b64decode(str(value).strip())
                        val_name = name.replace(".text", ".nodetypedvalue")
                        self.set(val_name, conv_val)
                    except Exception as e:
                        log.error("base64 conversion of '" + str(value) + "' failed. " + str(e))
                        
            except KeyError:
                pass

    def _strip_null_bytes(self, item):
        r = item
        if (isinstance(item, str)):
            r = item.replace("\x00", "")
        if (isinstance(item, list)):
            r = []
            for s in item:
                if (isinstance(s, str)):
                    r.append(s.replace("\x00", ""))
                else:
                    r.append(s)
        return r
                    
    def report_action(self, action, params=None, description=None, strip_null_bytes=False):

        # Strip out \x00 characters if needed.
        if (strip_null_bytes):
            action = self._strip_null_bytes(action)
            params = self._strip_null_bytes(params)
            description = self._strip_null_bytes(description)

        # Save the action for reporting.
        self.engine.report_action(action, params, description)

