VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIFceComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'-------------------------------'
' R Development Library 2.0 '
'-------------------------------'
'    R Interface Components '
'                   Version 1.0 '
'-------------------------------'
'                 Enum Declares '
'-------------------------------'
'Copyright © 1998-9 by R Software. All Rights Reserved


'Last Updated: 20/08/99

'20/08/99
'           - Added documentation for use with Document! VB

Option Explicit
DefInt A-Z

Public Enum ICTXClickStyles
 '##EN The <B>ICTXClickStyles</B> enum determines how the buttons on an <B>asxToolBox</B> react when clicked.
 txcsToolbar
 '##EM txcsToolbar (Default) Buttons stay down when pushed
 txcsNoStick
 '##EM txcsNoStick Buttons do not stay down when pushed
End Enum

Public Enum IFCTBBorderStyles
 '##EN The <B>IFCTBBorderStyles</B> enum determines how borders are drawn.
 tbbsNone
 '##EM tbbsNone No border
 tbbsInset
 '##EM tbbsInset A inset or lowered border
 tbbsRaised
 '##EM tbbsRaised (Default) A raised border
 tbbsFrame
 '##EM tbbsFrame A combination raised and lowered border, similar to the standard Frame control
 tbbsInsetButton
 '##EM tbbsInsetButton A double width lowered border
 tbbsRaisedButton
 '##EM tbbsRaisedButton A double width raised border
End Enum

Public Enum IFCTBButtonStyles
 '##EN The <B>IFCTBButtonStyles</B> enum determines how a button on an <B>asxToolbar</B> control is drawn and behaves
 tbbsPlaceholder
 '##EM tbbsPlaceholder A holder for another control - no visible button elements
 tbbsButton
 '##EM tbbsButton (Default) A standard clickable button
 tbbsSeparator
 '##EM tbbsSeparator A separator between buttons
 tbbsCheck
 '##EM tbbsSeparator A separator between buttons
 tbbsButtonGroup
 '##EM tbbsSeparator A separator between buttons

End Enum

Public Enum IFCTBStyle
 '##EN The <B>IFCTBStyle</B> enum determines how buttons on an <B>asxToolbar</B> control are sized
 tbtsVariable
 '##EM tbtsVariable (Default) Button size varies depending on caption text and picutre settings
 tbtsFixed
 '##EM tbtsFixed Button size is fixed according to object settings
End Enum

Public Enum IFCSBButtonStyles
 '##EN The <B>IFCSBButtonStyles</B> enum determines what system graphic is displayed on an <B>asxSysButton</B> control.
 sbbsCaptionClose
 '##EM sbbsCaptionClose Close window
 sbbsCaptionMinimise
 '##EM sbbsCaptionMinimise Mimimise window
 sbbsCaptionMaximise
 '##EM sbbsCaptionMaximise Maximise window
 sbbsCaptionRestore
 '##EM sbbsCaptionRestore Restore window
 sbbsCaptionHelp
 '##EM sbbsCaptionHelp What's this Help
 sbbsMenuArrow
 '##EM sbbsMenuArrow Submenu marker
 sbbsMenuCheck
 '##EM sbbsMenuCheck Checked item marker
 sbbsMenuBullet
 '##EM sbbsMenuBullet Selected item marker
 sbbsMenuArrowRight
 '##EM sbbsMenuArrowRight Submenu marker
 sbbsScrollUp
 '##EM sbbsScrollUp Up facing scroll
 sbbsScrollDown
 '##EM sbbsScrollDown Down facing scroll
 sbbsScrollLeft
 '##EM sbbsScrollLeft Left facing scroll
 sbbsScrollRight
 '##EM sbbsScrollRight Right facing scroll
 sbbsScrollComboBox
 '##EM sbbsScrollComboBox Combobox dropdown button
 sbbsScrollSizeGrip
 '##EM sbbsScrollSizeGrip Window size grip
 sbbsScrollSizeGripRight
 '##EM sbbsScrollSizeGripRight Window size grip
 sbbsButtonCheck
 '##EM sbbsButtonCheck Checkbox
 sbbsButtonRadio
 '##EM sbbsButtonRadio Option button
 sbbsButton3DState
 '##EM sbbsButton3DState Checkbox
 sbbsButtonPush
 '##EM sbbsButtonPush (Default) Command button
 sbbsNoGraphic
 '##EM sbbsNoGraphic No graphic
End Enum

Public Enum IFCSBButtonFlags
 '##EN The <B>IFCSBButtonFlags</B> enum determines additional behaviour and displayed flags for an <B>asxSysButton</B> control.
 sbbfNone = &H0
 '##EM sbbfNone (Default) No flags set
 sbbfInactive = &H100
 '##EM sbbfInactive Button is disabled
 sbbfPushed = &H200
 '##EM sbbfPushed Button is pushed (lowered)
 sbbfChecked = &H400
 '##EM sbbfChecked Button is checked
 sbbfAdjustrect = &H2000
 sbbfFlat = &H4000
 '##EM sbbfFlat Button is flat
 sbbfMono = &H8000
 '##EM sbbfMono Button is drawn in mono colours
End Enum

Public Enum IFCAlignments
 '##EN The <B>IFCAlignments</B> enum determines how text and graphics are positioned on a control.
 aiaTopLeft
 '##EM aiaTopLeft Align to the upper left corner of the control
 aiaTopCenter
 '##EM aiaTopCenter Align to the upper center of the control
 aiaTopRight
 '##EM aiaTopRight Align to the upper right corner of the control
 aiaCenterLeft
 '##EM aiaCenterLeft Align to the left of the control
 aiaCenterCenter
 '##EM aiaCenterCenter (Default) Align in the center of the control
 aiaCenterRight
 '##EM aiaCenterRight Align to the right of the control
 aiaBottomLeft
 '##EM aiaBottomLeft Align to the lower left corner of the control
 aiaBottomCenter
 '##EM aiaBottomCenter Align to the bottom center of the control
 aiaBottomRight
 '##EM aiaBottomRight Align to the lower left corner of the control
End Enum

Public Enum IFCFocusStyles
 '##EN The <B>IFCFocusStyles</B> enum determines how the focus rectangle is drawn on the active control.
 aifsAroundButton
 '##EM aifsAroundButton (Default) Focus rectangle is displayed around the button border
 aifsAroundText '##EM aifsAroundText Focus rectangle is displayed around caption text
 aifsNone
 '##EM aifsNone No focus rectangle is displayed
End Enum

Public Enum IFCAppearances
 '##EN The <B>IFCAppearances</B> enum determines how borders are drawn on a control
 ifcaStandard
 '##EM ifcaStandard (Default) Objects are drawn using the default 3D style
 ifcaFlat
 '##EM ifcaFlat Objects are drawn using a flat 3D style, generaly only when the mouse is hovered over the object
End Enum
 
'Public Enum IFCAppearancesEx
' ifaxStandard '(Default) Objects are drawn using the default 3D style
' ifaxFlat 'Objects are drawn using a flat 3D style, generaly only when the mouse is hovered over the object
' ifaxNoBorder 'No borders are drawn
'End Enum

Public Enum IFCStandardAlignments
 '##EN The <B>IFCStandardAlignments</B> enum determines how text is aligned to another object
 ifsaLeft
 '##EM ifsaLeft (Default) Drawn objects are aligned to the left
 ifsaRight
 '##EM ifsaRight Drawn objects are aligned to the right
 ifsaCenter
 '##EM ifsaCenter Drawn objects are center aligned
End Enum

Public Enum IFCStandardAlignmentsEx
 '##EN The <B>IFCStandardAlignmentsEx</B> enum determines how text is aligned to another object
 ifaxLeft
 '##EM ifaxLeft (Default) Drawn objects are aligned to the left
 ifaxRight
 '##EM ifaxRight Drawn objects are aligned to the right
 ifaxCenter
 '##EM ifaxCenter Drawn objects are center aligned
 ifaxJustify
 '##EM ifaxJustify Drawn objects are justified to fill the entire control width
End Enum

Public Enum IFCBorderStyles
 '##EN The <B>IFCBorderStyles</B> enum determines how borders are displayed in a control
 ifcbsNone
 '##EM ifcbsNone No border
 ifcbsFixedSingle
 '##EM ifcbsFixedSingle Fixed border
End Enum

Public Enum IFCOrientations
 '##EN The <B>IFCBorderStyles</B> enum determines how objects are drawn on a control
 ifcoHorizontal
 '##EM ifcoHorizontal (Default) Drawn objects are positioned horizontally.
 ifcoVertical
 '##EM ifcoVertical Drawn objects are positioned vertically.
End Enum

Public Enum IFCSBPanelModes
 '##EN The <B>IFCSBPanelModes</B> enum determines how a panel behaves in an <B>asxStatusBar</B> control
 sbpmNormal
 '##EM sbpmNormal (Default) Displayed text are drawn according to settings you provide
 sbpmCaps
 '##EM sbpmCaps Panel acts as a Caps Lock indicator
 sbpmNum
 '##EM sbpmNum Panel acts as a Num Lock indicator
 sbpmOver
 '##EM sbpmOver Panel acts as a Overtype indicator
 sppmScroll
 '##EM sppmScroll Panel acts as a Scroll Lock indicator
 eppmDateTime
 '##EM eppmDateTime Panel displays the date and time
End Enum

'Public Enum IFCSBSizeModes
' sbpsNormal '##EM (Default) Panel is size based on the width of the control versus the number of panels
' sbpsSpring '##EM Panel is sized to take up all available space
' sbpsAutoSize '##EM Panel is autosized depending on the contents of the panel
'End Enum

Public Enum IFCWhatPictures
 '##EN The <B>IFCWhatPictures</B> enum determines which pictures are used in a function or method.
 ifwpPicNormal
 '##EM ifwpPicNormal (Default) Images are displayed from the Picture property
 ifwpPicOver
 '##EM ifwpPicOver Images are displayed from the PictureOver property
 ifwpPicDown
 '##EM ifwpPicDown Images are displayed from the PictureDown property
End Enum

Public Enum IFCCaptionOptions
 iftoNoLabels
 '##EM iftoNoLabels No labels are displayed
 iftoShowLabels
 '##EM iftoShowLabels (Default) All labels are displayed
 iftoSelectedLabels
 '##EM iftoSelectedLabels Only certain labels are displayed
End Enum

Public Enum IFCCaptionAlignments
 ifcaCaptionOnTop
 '##EM ifcaCaptionOnTop Caption is displayed above button graphic
 ifcaCaptionOnBottom
 '##EM ifcaCaptionOnBottom (Default) Caption is displayed below button graphic
 ifcaCaptionOnLeft
 '##EM ifcaCaptionOnLeft Caption is displayed to left of button graphic
 ifcaCaptionOnRight
 '##EM ifcaCaptionOnRight Caption is displayed to the right of button graphic
End Enum

Public Enum IFCPBBorderStyles
 pbbsNone
 '##EM pbbsNone No borders are drawn
 pbbsWindows3x
 '##EM pbbsWindows3x Borders are drawn in Windows 3.1 style
 pbbsWindows95
 '##EM pbbsWindows95 Borders are drawn in standard 3D
 pbbsSingle
 '##EM pbbsSingle A single width border is drawn
 pbbsRound
 '##EM pbbsRound A round single width border is drawn
 pbbsSunken
 '##EM pbbsSunken Borders are drawn in standard 3D, with slightly different drawing for lowered buttons
End Enum

Public Enum IFCPBButtonModes
 pbbmNormal
 '##EM pbbmNormal (Default) Standard button operation
 pbbmClickable
 '##EM pbbmClickable Button remains pressed when first pushed, then raised with second push
 pbbmCheckedClickable
 '##EM pbbmCheckedClickable Same as <B>pbbmClickable</B> except that button is also checked when pushed
 pbbmDropDown
 '##EM pbbmDropDown Button displays a drop down menu button and splits into two seperate buttons
 pbbmDropDownNoSeparator
 '##EM pbbmDropDownNoSeparator Button displayed drop down arrow, but remains a single button
End Enum

Public Enum IFCPGStyles
 pgrsHorizontal
 '##EM pgrsHorizontal (Default) Arrows are positioned horizontally on each size of the control.
 pgrsHorizontalRight
 '##EM pgrsHorizontalRight Arrows are positioned horizontally on right right size of the control.
 pgrsVertical
 '##EM pgrsVertical Arrows are positioned vertically on the top and bottom of the control.
End Enum

'Public Enum IFCBSIconStyles
' bsisLarge
' bsisSmall
'End Enum

Public Enum IFCBackStyles
 ifbsTransparent
 '##EM ifbsTransparent Background color and any graphics are visible behind the control.
 ifbsOpaque
 '##EM ifbsOpaque (Default) The control's BackColor property setting fills the control and obscures any color or graphics behind it.
End Enum

Public Enum IFCDockStyles
 dckLocked = 0
 dckOffice97 = 1
 dckOfficeXP = 2
End Enum

