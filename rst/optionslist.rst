List Of Options
===============

All the options are in the ``vb2py.ini`` file located in the main ``vb2py`` folder.

General_

	IndentCharacter_

	IndentAmount_

	AttentionMarker_

	WarnAboutUnrenderedCode_

	LoadUserPlugins_

	LoggingLevel_

	DumpFormData_

	RespectPrivateStatus_

	PrivateDataPrefix_

	AlwaysUseRawStringLiterals_

	TryToExtractDocStrings_

	ReportPartialConversion_

	IncludeDebugCode_


Functions_

	ReturnVariableName_

	PreInitializeReturnVariable_

Select_

	EvaluateVariable_

	SelectVariablePrefix_

	UseNumericIndex_

Labels_

	IgnoreLabels_

With_

	`With EvaluateVariable`_

	WithVariablePrefix_

	`With UseNumericIndex`_

Properties_

	LetSetVariablePrefix_

	GetVariablePrefix_

Classes_

	UseNewStyleClasses_

General
-------

The list of all ``General`` options is shown in the following table::

	[General]
	# Space or Tab
	IndentCharacter = Space
	# Number of spaces/tabs
	IndentAmount = 4
	# Marker to use when code needs user attention
	AttentionMarker = VB2PY
	# Yes or No
	WarnAboutUnrenderedCode = Yes
	# Yes or No, whether to use user plugins or not. If No, system plugins will still work
	LoadUserPlugins = No
	# Default logging level, 0 is nothing
	LoggingLevel = 0
	# Yes or No, whether to dump form data to screen - Yes seems to crash the GUI!
	DumpFormData = No
	# Yes or No, whether the full VB parser is used to convert code
	UseFullParser = Yes
	# Yes or No, whether to respect Private status of variables
	RespectPrivateStatus = Yes
	# Prefix to use to tag data as private (Python normally uses __ but VB convention is m_)
	PrivateDataPrefix = __
	# Yes or No, whether to use raw strings for all literals - very safe but not necessarily good looking!
	AlwaysUseRawStringLiterals = No
	# Yes or No, whether to try to automatically extract docstrings from the code
	TryToExtractDocStrings = Yes
	# Yes or No, whether to return a partially converted file when an error is found
	ReportPartialConversion = Yes
	# Yes or No, whether to include debug code in the converted application
	IncludeDebugCode = No

IndentCharacter
~~~~~~~~~~~~~~~

Syntax: ``IndentCharacter = Space | Tab``

::

	# Space or Tab
	IndentCharacter = Space

Sets the indentation character as a Space or a Tab. The number of indent characters is set by the IndentAmount_.

The default is a space.

VB(VBCodeModule)::

    ' VB2PY-Set: General.IndentCharacter = Space
	If a = 10 Then
		b = 1
	Else
		b = 2
	End If
    ' VB2PY-Unset: General.IndentCharacter

Be careful when switching to ``Tab`` to set the IndentAmount_, or you will end up with four tabs!

VB(VBCodeModule)::

    ' VB2PY-Set: General.IndentCharacter = Tab
	If a = 10 Then
		b = 1
	Else
		b = 2
	End If
    ' VB2PY-Unset: General.IndentCharacter

IndentAmount
~~~~~~~~~~~~

Syntax: ``IndentAmount = <integer>``

::

	# Space or Tab
	IndentAmount = 4

Sets the number of IndentCharacter_ 's to be used to indent code blocks.

The default is 4.

VB(VBCodeModule)::

    ' VB2PY-Set: General.IndentAmount = 4
	If a = 10 Then
		b = 1
	Else
		b = 2
	End If
    ' VB2PY-Unset: General.IndentAmount

Other values are allowed.

VB(VBCodeModule)::

    ' VB2PY-Set: General.IndentAmount = 8
	If a = 10 Then
		b = 1
	Else
		b = 2
	End If
    ' VB2PY-Unset: General.IndentAmount

AttentionMarker
~~~~~~~~~~~~~~~

Syntax: ``AttentionMarker = <string>``

::

	# Marker to use when code needs user attention
	AttentionMarker = VB2PY

Sets the marker to use in comments when highlighting part of the converted code that needs attention.

The default is VB2PY.

VB(VBCodeModule)::

    ' VB2PY-Set: General.AttentionMarker = VB2PY
	On Error Goto 0
    ' VB2PY-Unset: General.AttentionMarker

Other values are allowed.

VB(VBCodeModule)::

    ' VB2PY-Set: General.AttentionMarker = TODO
	On Error Goto 0
    ' VB2PY-Unset: General.AttentionMarker

WarnAboutUnrenderedCode
~~~~~~~~~~~~~~~~~~~~~~~

Syntax: ``WarnAboutUnrenderedCode = Yes | No``

::

	# Yes or No
	WarnAboutUnrenderedCode = Yes

Determined whether an AttentionMarker_ is inserted in the Python code to highligh VB code which has not been rendered.

The default is Yes.

VB(VBCodeModule)::

    ' VB2PY-Set: General.WarnAboutUnrenderedCode = Yes
	On Error Goto 0
    ' VB2PY-Unset: General.WarnAboutUnrenderedCode

Other values are allowed.

VB(VBCodeModule)::

    ' VB2PY-Set: General.WarnAboutUnrenderedCode = No
	On Error Goto 0
    ' VB2PY-Unset: General.WarnAboutUnrenderedCode

LoadUserPlugins
~~~~~~~~~~~~~~~

Syntax: ``LoadUserPlugins = Yes | No``

::

	# Yes or No, whether to use user plugins or not. If No, system plugins will still work
	LoadUserPlugins = No


Determines whether user plug-ins are loaded and executed during the normal code conversion process. User plug-ins are kept in the ``extensions`` folder of the ``vb2py`` directory. System plug-ins are also located in this folder but are not affected by the value of this setting.

LoggingLevel
~~~~~~~~~~~~

Syntax: ``LoggingLevel = <integer>``

::

	# Default logging level, 0 is nothing
	LoggingLevel = 0


Sets the default logging level to use during the code conversion. If this is set to 0 then no logging messages will be output. The logging levels are defined in the standard Python logging module.

DumpFormData
~~~~~~~~~~~~

Syntax: ``DumpFormData = Yes | No``

::

	# Yes or No, whether to dump form data to screen - Yes seems to crash the GUI!
	DumpFormData = No


If this is set to ``Yes`` then the form classes will be dumped to the screen during form conversion. This may be useful if there is a problem during the conversion process.

The default is No.

RespectPrivateStatus
~~~~~~~~~~~~~~~~~~~~

Syntax: ``RespectPrivateStatus = Yes | No``

::

	# Yes or No, whether to respect Private status of variables
	RespectPrivateStatus = Yes

If this variable is set to ``Yes`` then variables, subroutines and functions in VB which are either explicitely or implicitely ``Private`` will have their Python names converted to have a PrivateDataPrefix_. Setting this variable to ``No`` will ignore the ``Private`` status of all variables.

The default is Yes.

VB(VBClassModule)::

    ' VB2PY-GlobalSet: General.RespectPrivateStatus = Yes
    Public Name As String
    Public Age As Single
    Private ID As Long

    Public Sub checkAge()
        If Age = 0 Then Age = 1
    End Sub
    '
    Private Sub setUp()
        ID = Rnd()
        If ID = 0 Then setUp
    End Sub
    ' VB2PY-Unset: General.RespectPrivateStatus

An example with No.

VB(VBClassModule)::

    ' VB2PY-GlobalSet: General.RespectPrivateStatus = No
    Public Name As String
    Public Age As Single
    Private ID As Long

    Public Sub checkAge()
        If Age = 0 Then Age = 1
    End Sub
    '
    Private Sub setUp()
        ID = Rnd()
        If ID = 0 Then setUp
    End Sub
    ' VB2PY-Unset: General.RespectPrivateStatus

PrivateDataPrefix
~~~~~~~~~~~~~~~~~

Syntax: ``PrivateDataPrefix = <string>``

::

	# Prefix to use to tag data as private (Python normally uses __ but VB convention is m_)
	PrivateDataPrefix = __

If RespectPrivateStatus_ is set to ``Yes`` then variables, subroutines and functions in VB which are either explicitely or implicitely ``Private`` will have their Python names converted to have a prefix and this setting determines what that prefix will be.

The default is ___.

VB(VBClassModule)::

    ' VB2PY-GlobalSet: General.PrivateDataPrefix = prv
    Public Name As String
    Public Age As Single
    Private ID As Long

    Public Sub checkAge()
        If Age = 0 Then Age = 1
    End Sub
    '
    Private Sub setUp()
        ID = Rnd()
        If ID = 0 Then setUp
    End Sub
    ' VB2PY-Unset: General.PrivateDataPrefix

If the value used is not "__" then the data will not be hidden as far as Python is concerned.

VB(VBClassModule)::

    ' VB2PY-GlobalSet: General.PrivateDataPrefix = m_
    Public Name As String
    Public Age As Single
    Private ID As Long

    Public Sub checkAge()
        If Age = 0 Then Age = 1
    End Sub
    '
    Private Sub setUp()
        ID = Rnd()
        If ID = 0 Then setUp
    End Sub
    ' VB2PY-Unset: General.PrivateDataPrefix

AlwaysUseRawStringLiterals
~~~~~~~~~~~~~~~~~~~~~~~~~~

Syntax: ``AlwaysUseRawStringLiterals = Yes | No``

::

	# Yes or No, whether to use raw strings for all literals - very safe but not necessarily good looking!
	AlwaysUseRawStringLiterals = No

By default, all VB strings are just converted to Python strings. However, if the VB string contains the backslash character then it is quite likely that the Python version will not be the same since Python will interpret the backslash as a control character. Setting the ``AlwaysUseRawStringLiterals`` option to ``Yes`` will cause all VB strings to be converted to raw Python strings (r"string"), which will prevent such problems.

The default is No.

VB(VBCodeModule)::

    ' VB2PY-GlobalSet: General.AlwaysUseRawStringLiterals = No
    myString = "a\path\name"
    ' VB2PY-Unset: General.AlwaysUseRawStringLiterals

Setting the option to ``Yes`` is safe but doesn't always look good in the code.

VB(VBCodeModule)::

    ' VB2PY-GlobalSet: General.AlwaysUseRawStringLiterals = Yes
    myString = "a\path\name"
    ' VB2PY-Unset: General.AlwaysUseRawStringLiterals

TryToExtractDocStrings
~~~~~~~~~~~~~~~~~~~~~~

Syntax: ``TryToExtractDocStrings = Yes | No``

::

	# Yes or No, whether to try to automatically extract docstrings from the code
	TryToExtractDocStrings = Yes

If ``TryToExtractDocStrings`` is set then any contiguous block of comment lines found at the start of a module are interpretted as a docstring and added to the class definition. The docstring terminates with the first non-comment line.

The default is ``No``.

VB(VBCodeModule)::

    ' VB2PY-GlobalSet: General.TryToExtractDocStrings = No
	' This is the documentation for the module
	' This line is also documentation
	' So is this one
	' And this is the last

    Public Name As String
    Public Age As Single
    Private ID As Long

    Public Sub checkAge()
        If Age = 0 Then Age = 1
    End Sub
    '
    Private Sub setUp()
        ID = Rnd()
        If ID = 0 Then setUp
    End Sub
    ' VB2PY-Unset: General.TryToExtractDocStrings

When the option is ``Yes`` docstrings will be created.

VB(VBCodeModule)::

    ' VB2PY-GlobalSet: General.TryToExtractDocStrings = Yes
	' This is the documentation for the module
	' This line is also documentation
	' So is this one
	' And this is the last

    Public Name As String
    Public Age As Single
    Private ID As Long

    Public Sub checkAge()
        If Age = 0 Then Age = 1
    End Sub
    '
    Private Sub setUp()
        ID = Rnd()
        If ID = 0 Then setUp
    End Sub
    ' VB2PY-Unset: General.TryToExtractDocStrings

ReportPartialConversion
~~~~~~~~~~~~~~~~~~~~~~~

Syntax: ``ReportPartialConversion = Yes | No``

::

	# Yes or No, whether to return a partially converted file when an error is found
	ReportPartialConversion = Yes

This option is used to determine what happens when the conversion fails for some reason. If the option is set to ``Yes`` then the conversion will return as much code as it can. If the option is set to ``No`` then the conversion will just fail and return nothing at all.

The default is ``Yes``.

VB(VBCodeModule)::

    ' VB2PY-Set: General.ReportPartialConversion = Yes
	a = 10
	b = 20
	c = 30
	something that wont convert
	d = 40
	e = 50
   ' VB2PY-Unset: General.ReportPartialConversion

When the option is ``No`` you wont get any output if there is an error.

VB(VBCodeModule)::

    ' VB2PY-Set: General.ReportPartialConversion = No
	a = 10
	b = 20
	c = 30
	something that wont convert
	d = 40
	e = 50
   ' VB2PY-Unset: General.ReportPartialConversion

IncludeDebugCode
~~~~~~~~~~~~~~~~

Syntax: ``IncludeDebugCode = Yes | No``

::

	# Yes or No, whether to include debug code in the converted application
	IncludeDebugCode = No

This option is used to determine whether debug code is included in the converted application. If the option is ``Yes`` then a ``from vbdebug import *`` will be inserted at the top of each module. ``vbdebug`` includes code to access the logger and is required if you need to view the output from ``Debug.Print`` statements.

The default is ``No``.

VB(VBCodeModule)::

    ' VB2PY-Set: General.IncludeDebugCode = No
	a = 10
	b = 20
	c = 30
   ' VB2PY-Unset: General.IncludeDebugCode

When the option is ``Yes`` you get the extra ``import`` statement

VB(VBCodeModule)::

    ' VB2PY-Set: General.IncludeDebugCode = Yes
	a = 10
	b = 20
	c = 30
   ' VB2PY-Unset: General.IncludeDebugCode

Functions
---------

The list of all ``Function`` options is shown in the following table::

	[Functions]
	# Name of variable used in Functions
	ReturnVariableName = _ret
	# Yes or No, leave at Yes unless good reasons!
	PreInitializeReturnVariable = Yes


ReturnVariableName
~~~~~~~~~~~~~~~~~~

Syntax: ``ReturnVariableName = <string>``

::

	# Name of variable used in Functions
	ReturnVariableName = _ret

This option allows the return variable name to be specified. No checking is done to ensure that the name does not clash with local or global variables, so care should be taken when selecting a suitable name.

VB::

    Dim moduleGlobal1, moduleGlobal2

    ' VB2PY-GlobalSet: Functions.ReturnVariableName = _MyFunc
    Function MyFunc(X, Optional Y, Optional Z=20)
        Dim subLocal
        subLocal = X + Y + Z + moduleGlobal
        moduleGlobal2 = moduleGlobal2 + 1
        MyFunc = subLocal*10
    End Function
    ' VB2PY-Unset: Functions.ReturnVariableName

    a = MyFunc(1, 2)
    a = MyFunc(1, Z:=10)


PreInitializeReturnVariable
~~~~~~~~~~~~~~~~~~~~~~~~~~~

Syntax: ``PreInitializeReturnVariable = Yes | No``

::

	# Yes or No, leave at Yes unless good reasons!
	PreInitializeReturnVariable = Yes

By default the return variable is initialized to ``None`` at the start of the function so that an error does not occur in the event that the function returns before the return variable has been assigned to. This option allows this initialization step to be omitted and is safe as long as all return paths from the function include an explicit assignment to the return value variable.

VB::

    Dim moduleGlobal1, moduleGlobal2

    ' VB2PY-GlobalSet: Functions.PreInitializeReturnVariable = Yes
    Function MyFunc(X, Optional Y, Optional Z=20)
        Dim subLocal
        subLocal = X + Y + Z + moduleGlobal
        moduleGlobal2 = moduleGlobal2 + 1
        MyFunc = subLocal*10
    End Function
    ' VB2PY-Unset: Functions.PreInitializeReturnVariable

    a = MyFunc(1, 2)
    a = MyFunc(1, Z:=10)

Compare this with,

VB::

    Dim moduleGlobal1, moduleGlobal2

    ' VB2PY-GlobalSet: Functions.PreInitializeReturnVariable = No
    Function MyFunc(X, Optional Y, Optional Z=20)
        Dim subLocal
        subLocal = X + Y + Z + moduleGlobal
        moduleGlobal2 = moduleGlobal2 + 1
        MyFunc = subLocal*10
    End Function
    ' VB2PY-Unset: Functions.PreInitializeReturnVariable

    a = MyFunc(1, 2)
    a = MyFunc(1, Z:=10)

Select
------

The list of all ``Select`` options is shown in the following table::

	[Select]
	# Once or EachTime, how many times to evaluate the case variable
	EvaluateVariable = Once
	# Name of select variable (only used if EvaluateVariable is Once)
	SelectVariablePrefix = _select
	# Yes or No, use numeric index on select variable (needed if you every have nested Selects and EvaluateVariable = Once)
	UseNumericIndex = Yes

EvaluateVariable
~~~~~~~~~~~~~~~~

Syntax: ``EvaluateVariable = Yes | No``

::

	# Once or EachTime, how many times to evaluate the case variable
	EvaluateVariable = Once

The default behaviour when converting a ``Select`` is to evaluate the select expression once at the start of the block. By setting this option to ``EachTime`` you can force the expression to be evaluated for each ``if/elif`` statement. This generally looks cleaner but can lead to undesired side effects or slow run times depending on how expensive [1]_ the expression is to calculate.

VB::

    ' VB2PY-Set: Select.EvaluateVariable = Once
    Select Case Value
        Case 1
            DoOne
        Case 2
            DoTwo
        Case 3, 4
            DoThreeOrFour
        Case 5 To 10
            DoFiveToTen
        Case Else
            DoElse
    End Select
    ' VB2PY-Unset: Select.EvaluateVariable

Compare this to,

VB::

    ' VB2PY-Set: Select.EvaluateVariable = EachTime
    Select Case Value
        Case 1
            DoOne
        Case 2
            DoTwo
        Case 3, 4
            DoThreeOrFour
        Case 5 To 10
            DoFiveToTen
        Case Else
            DoElse
    End Select
    ' VB2PY-Unset: Select.EvaluateVariable

SelectVariablePrefix
~~~~~~~~~~~~~~~~~~~~

Syntax: ``SelectVariablePrefix = <string>``

::

	# Name of select variable (only used if EvaluateVariable is Once)
	SelectVariablePrefix = _select

When EvaluateVariable_ is set to ``Once``, this option determines the prefix used to name the variable used in the select. If UseNumericIndex_ is set to ``No`` then this option sets the variable name used, otherwise this is the prefix and the final variable will also include a unique ID number.

VB::

    ' VB2PY-Set: Select.SelectVariablePrefix = selectVariable
    Select Case Value
        Case 1
            DoOne
        Case 2
            DoTwo
        Case 3, 4
            DoThreeOrFour
        Case 5 To 10
            DoFiveToTen
        Case Else
            DoElse
    End Select
    ' VB2PY-Unset: Select.SelectVariablePrefix

UseNumericIndex
~~~~~~~~~~~~~~~

Syntax: ``UseNumericIndex = Yes | No``

::

	# Yes or No, use numeric index on select variable (needed if you every have nested Selects and EvaluateVariable = Once)
	UseNumericIndex = Yes

When EvaluateVariable_ is set to ``Once``, this option determines whether a unique ID number is appended to the SelectVariablePrefix_ to determine the variable name used to hold the select expression. If used, the index is incremented for each ``select`` constuct found. This option is always required to be ``Yes`` where the code includes nested ``Select`` blocks *and* EvaluateVariable_ is set to ``Once``. If neither of these conditions applies then it is safe to set this to ``No``

VB::

    ' VB2PY-Set: Select.UseNumericIndex = Yes
    Select Case Value
        Case 1
            DoOne
        Case 2
            DoTwo
        Case 3, 4
            DoThreeOrFour
        Case 5 To 10
            DoFiveToTen
        Case Else
            DoElse
    End Select
    ' VB2PY-Unset: Select.UseNumericIndex

Comapre this to,

VB::

    ' VB2PY-Set: Select.UseNumericIndex = No
    Select Case Value
        Case 1
            DoOne
        Case 2
            DoTwo
        Case 3, 4
            DoThreeOrFour
        Case 5 To 10
            DoFiveToTen
        Case Else
            DoElse
    End Select
    ' VB2PY-Unset: Select.UseNumericIndex

Labels
------

The list of all ``Labels`` options is shown in the following table::

	[Labels]
	# Yes or No, ignore labels completely
	IgnoreLabels = Yes

IgnoreLabels
~~~~~~~~~~~~

Syntax: ``IgnoreLabels = Yes | No``

::

	# Yes or No, ignore labels completely
	IgnoreLabels = Yes

Labels are not supported in vb2Py v0.2. If you have VB code with labels on every line then you will get a huge number of attention markers telling you that the label was not converted. You can silence these warning by setting the ``IgnoreLabels`` option to ``Yes``.

VB::

    ' VB2PY-Set: Labels.IgnoreLabels = No
	10: a=1
	20: b=2
	30: c=3
    ' VB2PY-Unset: Labels.IgnoreLabels

Comapre this to,

VB::

    ' VB2PY-Set: Labels.IgnoreLabels = Yes
	10: a=1
	20: b=2
	30: c=3
    ' VB2PY-Unset: Labels.IgnoreLabels

With
----

The list of all ``With`` options is shown in the following table::

	[With]
	# Once or EachTime, how many times to evaluate the with variable
	EvaluateVariable = Once
	# Name of with variable (only used if EvaluateVariable is Once)
	WithVariablePrefix = _with
	# Yes or No, use numeric index on with variable (needed if you every have nested Withs and EvaluateVariable = Once)
	UseNumericIndex = Yes

With EvaluateVariable
~~~~~~~~~~~~~~~~~~~~~

Syntax: ``EvaluateVariable = Yes | No``

::

	[With]
	# Once or EachTime, how many times to evaluate the with variable

The default behaviour is to evaluate the ``With`` object once at the start of the block. By setting this option to ``EachTime`` you can force the object to be evaluated each time it is required. This generally looks more natural but can lead to undesired side effects or slow run times depending on how expensive [1]_ the object is to calculate.

VB::

    ' VB2PY-Set: With.EvaluateVariable = Once
    With MyObject
        .Height = 10
        .Width = .Height * .ScaleFactor
    End With
    ' VB2PY-Unset: With.EvaluateVariable

Compare this to,

VB::

    ' VB2PY-Set: With.EvaluateVariable = EveryTime
    With MyObject
        .Height = 10
        .Width = .Height * .ScaleFactor
    End With
    ' VB2PY-Unset: With.EvaluateVariable


WithVariablePrefix
~~~~~~~~~~~~~~~~~~

Syntax: ``WithVariablePrefix = <string>``

::

	# Name of with variable (only used if EvaluateVariable is Once)
	WithVariablePrefix = _select

When `With EvaluateVariable`_ is set to ``Once``, this option determines the prefix used to name the variable used in the ``With``. If `With UseNumericIndex`_ is set to ``No`` then this option sets the variable name used, otherwise this is the prefix and the final variable will also include a unique ID number.

VB::

    ' VB2PY-Set: With.WithVariablePrefix = withVariable
    With MyObject
        .Height = 10
        .Width = .Height * .ScaleFactor
    End With
    ' VB2PY-Unset: With.WithVariablePrefix

With UseNumericIndex
~~~~~~~~~~~~~~~~~~~~

Syntax: ``UseNumericIndex = Yes | No``

::

	# Yes or No, use numeric index on select variable (needed if you every have nested Selects and EvaluateVariable = Once)
	UseNumericIndex = Yes

When `With EvaluateVariable`_ is set to ``Once``, this option determines whether a unique ID number is appended to the WithVariablePrefix_ to determine the variable name used to hold the object. If used, the index is incremented for each ``With`` constuct found. This option is always required to be ``Yes`` where the code includes nested ``With`` blocks *and* `With EvaluateVariable`_ is set to ``Once``. If neither of these conditions applies then it is safe to set this to ``No``

VB::

    ' VB2PY-Set: With.UseNumericIndex = No
    With MyObject
        .Height = 10
        .Width = .Height * .ScaleFactor
    End With
    ' VB2PY-Unset: With.UseNumericIndex

Compare this to,

VB::

    ' VB2PY-Set: With.UseNumericIndex = Yes
    With MyObject
        .Height = 10
        .Width = .Height * .ScaleFactor
    End With
    ' VB2PY-Unset: With.UseNumericIndex

Properties
----------

The list of all ``Property`` options is shown in the following table::

	[Properties]
	# Prefix to add to property Let/Set function name
	LetSetVariablePrefix = set
	# Prefix to add to property Get function name
	GetVariablePrefix = get

LetSetVariablePrefix
~~~~~~~~~~~~~~~~~~~~

Syntax: ``LetSetVariablePrefix = <string>``

::

	# Prefix to add to property Let/Set function name
	LetSetVariablePrefix = set

In class modules where properties are defined, vb2Py creates ``get`` and ``set`` methods to access and assign to the property. Since VB uses a syntactic form to distinguish between the getters and setters but Python uses different names with the same syntax there is a need to automatically generate a name for the ``get`` and ``set`` methods. The ``getter`` and ``setter`` methods are determined by the LetSetVariablePrefix_ and GetVariablePrefix_ respectively.

VB::

    ' VB2PY-Set: Properties.LetSetVariablePrefix = doSet_
    Dim mName As String
    Dim mAge As Single

    Public Property Let Name(Value)
        mName = Value
    End Property
    '
    Public Property Get Name()
        Name = mName
    End Property

    ' VB2PY-Unset: Properties.LetSetVariablePrefix

GetVariablePrefix
~~~~~~~~~~~~~~~~~

Syntax: ``GetVariablePrefix = <string>``

::

	# Prefix to add to property Get function name
	GetVariablePrefix = set

In class modules where properties are defined, vb2Py creates ``get`` and ``set`` methods to access and assign to the property. Since VB uses a syntactic form to distinguish between the getters and setters but Python uses different names with the same syntax there is a need to automatically generate a name for the ``get`` and ``set`` methods. The ``getter`` and ``setter`` methods are determined by the LetSetVariablePrefix_ and GetVariablePrefix_ respectively.

VB::

    ' VB2PY-Set: Properties.GetVariablePrefix = doGet_
    Dim mName As String
    Dim mAge As Single

    Public Property Let Name(Value)
        mName = Value
    End Property
    '
    Public Property Get Name()
        Name = mName
    End Property

    ' VB2PY-Unset: Properties.GetVariablePrefix

Classes
-------

The list of all ``Class`` options is shown in the following table::

	[Classes]
	# Yes or No, whether to use new style classes for all classes
	UseNewStyleClasses = Yes

UseNewStyleClasses
~~~~~~~~~~~~~~~~~~

Syntax: ``UseNewStyleClasses = Yes | No``

::

	# Yes or No, whether to use new style classes for all classes
	UseNewStyleClasses = Yes

By default, all classes are created as *new style* Python classes (inheriting from ``Object``). Old style classes can be created by setting the ``UseNewStyleClasses`` option to ``No``.

VB(VBClassModule)::

    ' VB2PY-GlobalSet: Classes.UseNewStyleClasses = Yes
    Public Name As String
    Public Age As Single
    Private ID As Long

    Public Sub checkAge()
        If Age = 0 Then Age = 1
    End Sub
    '
    Private Sub setUp()
        ID = Rnd()
        If ID = 0 Then setUp
    End Sub
    ' VB2PY-Unset: Classes.UseNewStyleClasses

Compare this to,

VB(VBClassModule)::

    ' VB2PY-GlobalSet: Classes.UseNewStyleClasses = No
    Public Name As String
    Public Age As Single
    Private ID As Long

    Public Sub checkAge()
        If Age = 0 Then Age = 1
    End Sub
    '
    Private Sub setUp()
        ID = Rnd()
        If ID = 0 Then setUp
    End Sub
    ' VB2PY-Unset: Classes.UseNewStyleClasses

--------------

.. [1] Expensive as in CPU time.
