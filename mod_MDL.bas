Attribute VB_Name = "mod_MDL"
' ####################################################################################################################
' #
' #  MDL Model module - Copyright (c) TPD Software
' #
' #     This module contains the list of precalculated normals and errors returned by the classes
' #
' ####################################################################################################################

Public Enum MDL_ERRORS
   MDL_OK = 0
   MDL_INVALID_ID = 1
   MDL_INVALID_VERSION = 2
   MDL_LOAD_ERROR = 3
   MDL_MISSING_SKIN = 4
   MDL_CUSTOM_SKIN_OVERWRITTEN = 5
End Enum

Public Normals(161, 2) As Single


' ## List of precalculated normals
Sub LoadNormals()
    Normals(0, 0) = -0.525731:   Normals(0, 1) = 0#:          Normals(0, 2) = 0.850651
    Normals(1, 0) = -0.442863:   Normals(1, 1) = 0.238856:    Normals(1, 2) = 0.864188
    Normals(2, 0) = -0.295242:   Normals(2, 1) = 0#:          Normals(2, 2) = 0.955423
    Normals(3, 0) = -0.309017:   Normals(3, 1) = 0.5:         Normals(3, 2) = 0.809017
    Normals(4, 0) = -0.16246:    Normals(4, 1) = 0.262866:    Normals(4, 2) = 0.951056
    Normals(5, 0) = 0#:          Normals(5, 1) = 0#:          Normals(5, 2) = 1#
    Normals(6, 0) = 0#:          Normals(6, 1) = 0.850651:    Normals(6, 2) = 0.525731
    Normals(7, 0) = -0.147621:   Normals(7, 1) = 0.716567:    Normals(7, 2) = 0.681718
    Normals(8, 0) = 0.147621:    Normals(8, 1) = 0.716567:    Normals(8, 2) = 0.681718
    Normals(9, 0) = 0#:          Normals(9, 1) = 0.525731:    Normals(9, 2) = 0.850651
    Normals(10, 0) = 0.309017:   Normals(10, 1) = 0.5:        Normals(10, 2) = 0.809017
    Normals(11, 0) = 0.525731:   Normals(11, 1) = 0#:         Normals(11, 2) = 0.850651
    Normals(12, 0) = 0.295242:   Normals(12, 1) = 0#:         Normals(12, 2) = 0.955423
    Normals(13, 0) = 0.442863:   Normals(13, 1) = 0.238856:   Normals(13, 2) = 0.864188
    Normals(14, 0) = 0.16246:    Normals(14, 1) = 0.262866:   Normals(14, 2) = 0.951056
    Normals(15, 0) = -0.681718:  Normals(15, 1) = 0.147621:   Normals(15, 2) = 0.716567
    Normals(16, 0) = -0.809017:  Normals(16, 1) = 0.309017:   Normals(16, 2) = 0.5
    Normals(17, 0) = -0.587785:  Normals(17, 1) = 0.425325:   Normals(17, 2) = 0.688191
    Normals(18, 0) = -0.850651:  Normals(18, 1) = 0.525731:   Normals(18, 2) = 0#
    Normals(19, 0) = -0.864188:  Normals(19, 1) = 0.442863:   Normals(19, 2) = 0.238856
    Normals(20, 0) = -0.716567:  Normals(20, 1) = 0.681718:   Normals(20, 2) = 0.147621
    Normals(21, 0) = -0.688191:  Normals(21, 1) = 0.587785:   Normals(21, 2) = 0.425325
    Normals(22, 0) = -0.5:       Normals(22, 1) = 0.809017:   Normals(22, 2) = 0.309017
    Normals(23, 0) = -0.238856:  Normals(23, 1) = 0.864188:   Normals(23, 2) = 0.442863
    Normals(24, 0) = -0.425325:  Normals(24, 1) = 0.688191:   Normals(24, 2) = 0.587785
    Normals(25, 0) = -0.716567:  Normals(25, 1) = 0.681718:   Normals(25, 2) = -0.147621
    Normals(26, 0) = -0.5:       Normals(26, 1) = 0.809017:   Normals(26, 2) = -0.309017
    Normals(27, 0) = -0.525731:  Normals(27, 1) = 0.850651:   Normals(27, 2) = 0#
    Normals(28, 0) = 0#:         Normals(28, 1) = 0.850651:   Normals(28, 2) = -0.525731
    Normals(29, 0) = -0.238856:  Normals(29, 1) = 0.864188:   Normals(29, 2) = -0.442863
    Normals(30, 0) = 0#:         Normals(30, 1) = 0.955423:   Normals(30, 2) = -0.295242
    Normals(31, 0) = -0.262866:  Normals(31, 1) = 0.951056:   Normals(31, 2) = -0.16246
    Normals(32, 0) = 0#:         Normals(32, 1) = 1#:         Normals(32, 2) = 0#
    Normals(33, 0) = 0#:         Normals(33, 1) = 0.955423:   Normals(33, 2) = 0.295242
    Normals(34, 0) = -0.262866:  Normals(34, 1) = 0.951056:   Normals(34, 2) = 0.16246
    Normals(35, 0) = 0.238856:   Normals(35, 1) = 0.864188:   Normals(35, 2) = 0.442863
    Normals(36, 0) = 0.262866:   Normals(36, 1) = 0.951056:   Normals(36, 2) = 0.16246
    Normals(37, 0) = 0.5:        Normals(37, 1) = 0.809017:   Normals(37, 2) = 0.309017
    Normals(38, 0) = 0.238856:   Normals(38, 1) = 0.864188:   Normals(38, 2) = -0.442863
    Normals(39, 0) = 0.262866:   Normals(39, 1) = 0.951056:   Normals(39, 2) = -0.16246
    Normals(40, 0) = 0.5:        Normals(40, 1) = 0.809017:   Normals(40, 2) = -0.309017
    Normals(41, 0) = 0.850651:   Normals(41, 1) = 0.525731:   Normals(41, 2) = 0#
    Normals(42, 0) = 0.716567:   Normals(42, 1) = 0.681718:   Normals(42, 2) = 0.147621
    Normals(43, 0) = 0.716567:   Normals(43, 1) = 0.681718:   Normals(43, 2) = -0.147621
    Normals(44, 0) = 0.525731:   Normals(44, 1) = 0.850651:   Normals(44, 2) = 0#
    Normals(45, 0) = 0.425325:   Normals(45, 1) = 0.688191:   Normals(45, 2) = 0.587785
    Normals(46, 0) = 0.864188:   Normals(46, 1) = 0.442863:   Normals(46, 2) = 0.238856
    Normals(47, 0) = 0.688191:   Normals(47, 1) = 0.587785:   Normals(47, 2) = 0.425325
    Normals(48, 0) = 0.809017:   Normals(48, 1) = 0.309017:   Normals(48, 2) = 0.5
    Normals(49, 0) = 0.681718:   Normals(49, 1) = 0.147621:   Normals(49, 2) = 0.716567
    Normals(50, 0) = 0.587785:   Normals(50, 1) = 0.425325:   Normals(50, 2) = 0.688191
    Normals(51, 0) = 0.955423:   Normals(51, 1) = 0.295242:   Normals(51, 2) = 0#
    Normals(52, 0) = 1#:         Normals(52, 1) = 0#:         Normals(52, 2) = 0#
    Normals(53, 0) = 0.951056:   Normals(53, 1) = 0.16246:    Normals(53, 2) = 0.262866
    Normals(54, 0) = 0.850651:   Normals(54, 1) = -0.525731:  Normals(54, 2) = 0#
    Normals(55, 0) = 0.955423:   Normals(55, 1) = -0.295242:  Normals(55, 2) = 0#
    Normals(56, 0) = 0.864188:   Normals(56, 1) = -0.442863:  Normals(56, 2) = 0.238856
    Normals(57, 0) = 0.951056:   Normals(57, 1) = -0.16246:   Normals(57, 2) = 0.262866
    Normals(58, 0) = 0.809017:   Normals(58, 1) = -0.309017:  Normals(58, 2) = 0.5
    Normals(59, 0) = 0.681718:   Normals(59, 1) = -0.147621:  Normals(59, 2) = 0.716567
    Normals(60, 0) = 0.850651:   Normals(60, 1) = 0#:         Normals(60, 2) = 0.525731
    Normals(61, 0) = 0.864188:   Normals(61, 1) = 0.442863:   Normals(61, 2) = -0.238856
    Normals(62, 0) = 0.809017:   Normals(62, 1) = 0.309017:   Normals(62, 2) = -0.5
    Normals(63, 0) = 0.951056:   Normals(63, 1) = 0.16246:    Normals(63, 2) = -0.262866
    Normals(64, 0) = 0.525731:   Normals(64, 1) = 0#:         Normals(64, 2) = -0.850651
    Normals(65, 0) = 0.681718:   Normals(65, 1) = 0.147621:   Normals(65, 2) = -0.716567
    Normals(66, 0) = 0.681718:   Normals(66, 1) = -0.147621:  Normals(66, 2) = -0.716567
    Normals(67, 0) = 0.850651:   Normals(67, 1) = 0#:         Normals(67, 2) = -0.525731
    Normals(68, 0) = 0.809017:   Normals(68, 1) = -0.309017:  Normals(68, 2) = -0.5
    Normals(69, 0) = 0.864188:   Normals(69, 1) = -0.442863:  Normals(69, 2) = -0.238856
    Normals(70, 0) = 0.951056:   Normals(70, 1) = -0.16246:   Normals(70, 2) = -0.262866
    Normals(71, 0) = 0.147621:   Normals(71, 1) = 0.716567:   Normals(71, 2) = -0.681718
    Normals(72, 0) = 0.309017:   Normals(72, 1) = 0.5:        Normals(72, 2) = -0.809017
    Normals(73, 0) = 0.425325:   Normals(73, 1) = 0.688191:   Normals(73, 2) = -0.587785
    Normals(74, 0) = 0.442863:   Normals(74, 1) = 0.238856:   Normals(74, 2) = -0.864188
    Normals(75, 0) = 0.587785:   Normals(75, 1) = 0.425325:   Normals(75, 2) = -0.688191
    Normals(76, 0) = 0.688191:   Normals(76, 1) = 0.587785:   Normals(76, 2) = -0.425325
    Normals(77, 0) = -0.147621:  Normals(77, 1) = 0.716567:   Normals(77, 2) = -0.681718
    Normals(78, 0) = -0.309017:  Normals(78, 1) = 0.5:        Normals(78, 2) = -0.809017
    Normals(79, 0) = 0#:         Normals(79, 1) = 0.525731:   Normals(79, 2) = -0.850651
    Normals(80, 0) = -0.525731:  Normals(80, 1) = 0#:         Normals(80, 2) = -0.850651
    Normals(81, 0) = -0.442863:  Normals(81, 1) = 0.238856:   Normals(81, 2) = -0.864188
    Normals(82, 0) = -0.295242:  Normals(82, 1) = 0#:         Normals(82, 2) = -0.955423
    Normals(83, 0) = -0.16246:   Normals(83, 1) = 0.262866:   Normals(83, 2) = -0.951056
    Normals(84, 0) = 0#:         Normals(84, 1) = 0#:         Normals(84, 2) = -1#
    Normals(85, 0) = 0.295242:   Normals(85, 1) = 0#:         Normals(85, 2) = -0.955423
    Normals(86, 0) = 0.16246:    Normals(86, 1) = 0.262866:   Normals(86, 2) = -0.951056
    Normals(87, 0) = -0.442863:  Normals(87, 1) = -0.238856:  Normals(87, 2) = -0.864188
    Normals(88, 0) = -0.309017:  Normals(88, 1) = -0.5:       Normals(88, 2) = -0.809017
    Normals(89, 0) = -0.16246:   Normals(89, 1) = -0.262866:  Normals(89, 2) = -0.951056
    Normals(90, 0) = 0#:         Normals(90, 1) = -0.850651:  Normals(90, 2) = -0.525731
    Normals(91, 0) = -0.147621:  Normals(91, 1) = -0.716567:  Normals(91, 2) = -0.681718
    Normals(92, 0) = 0.147621:   Normals(92, 1) = -0.716567:  Normals(92, 2) = -0.681718
    Normals(93, 0) = 0#:         Normals(93, 1) = -0.525731:  Normals(93, 2) = -0.850651
    Normals(94, 0) = 0.309017:   Normals(94, 1) = -0.5:       Normals(94, 2) = -0.809017
    Normals(95, 0) = 0.442863:   Normals(95, 1) = -0.238856:  Normals(95, 2) = -0.864188
    Normals(96, 0) = 0.16246:    Normals(96, 1) = -0.262866:  Normals(96, 2) = -0.951056
    Normals(97, 0) = 0.238856:   Normals(97, 1) = -0.864188:  Normals(97, 2) = -0.442863
    Normals(98, 0) = 0.5:        Normals(98, 1) = -0.809017:  Normals(98, 2) = -0.309017
    Normals(99, 0) = 0.425325:   Normals(99, 1) = -0.688191:  Normals(99, 2) = -0.587785
    Normals(100, 0) = 0.716567:  Normals(100, 1) = -0.681718: Normals(100, 2) = -0.147621
    Normals(101, 0) = 0.688191:  Normals(101, 1) = -0.587785: Normals(101, 2) = -0.425325
    Normals(102, 0) = 0.587785:  Normals(102, 1) = -0.425325: Normals(102, 2) = -0.688191
    Normals(103, 0) = 0#:        Normals(103, 1) = -0.955423: Normals(103, 2) = -0.295242
    Normals(104, 0) = 0#:        Normals(104, 1) = -1#:       Normals(104, 2) = 0#
    Normals(105, 0) = 0.262866:  Normals(105, 1) = -0.951056: Normals(105, 2) = -0.16246
    Normals(106, 0) = 0#:        Normals(106, 1) = -0.850651: Normals(106, 2) = 0.525731
    Normals(107, 0) = 0#:        Normals(107, 1) = -0.955423: Normals(107, 2) = 0.295242
    Normals(108, 0) = 0.238856:  Normals(108, 1) = -0.864188: Normals(108, 2) = 0.442863
    Normals(109, 0) = 0.262866:  Normals(109, 1) = -0.951056: Normals(109, 2) = 0.16246
    Normals(110, 0) = 0.5:       Normals(110, 1) = -0.809017: Normals(110, 2) = 0.309017
    Normals(111, 0) = 0.716567:  Normals(111, 1) = -0.681718: Normals(111, 2) = 0.147621
    Normals(112, 0) = 0.525731:  Normals(112, 1) = -0.850651: Normals(112, 2) = 0#
    Normals(113, 0) = -0.238856: Normals(113, 1) = -0.864188: Normals(113, 2) = -0.442863
    Normals(114, 0) = -0.5:      Normals(114, 1) = -0.809017: Normals(114, 2) = -0.309017
    Normals(115, 0) = -0.262866: Normals(115, 1) = -0.951056: Normals(115, 2) = -0.16246
    Normals(116, 0) = -0.850651: Normals(116, 1) = -0.525731: Normals(116, 2) = 0#
    Normals(117, 0) = -0.716567: Normals(117, 1) = -0.681718: Normals(117, 2) = -0.147621
    Normals(118, 0) = -0.716567: Normals(118, 1) = -0.681718: Normals(118, 2) = 0.147621
    Normals(119, 0) = -0.525731: Normals(119, 1) = -0.850651: Normals(119, 2) = 0#
    Normals(120, 0) = -0.5:      Normals(120, 1) = -0.809017: Normals(120, 2) = 0.309017
    Normals(121, 0) = -0.238856: Normals(121, 1) = -0.864188: Normals(121, 2) = 0.442863
    Normals(122, 0) = -0.262866: Normals(122, 1) = -0.951056: Normals(122, 2) = 0.16246
    Normals(123, 0) = -0.864188: Normals(123, 1) = -0.442863: Normals(123, 2) = 0.238856
    Normals(124, 0) = -0.809017: Normals(124, 1) = -0.309017: Normals(124, 2) = 0.5
    Normals(125, 0) = -0.688191: Normals(125, 1) = -0.587785: Normals(125, 2) = 0.425325
    Normals(126, 0) = -0.681718: Normals(126, 1) = -0.147621: Normals(126, 2) = 0.716567
    Normals(127, 0) = -0.442863: Normals(127, 1) = -0.238856: Normals(127, 2) = 0.864188
    Normals(128, 0) = -0.587785: Normals(128, 1) = -0.425325: Normals(128, 2) = 0.688191
    Normals(129, 0) = -0.309017: Normals(129, 1) = -0.5:      Normals(129, 2) = 0.809017
    Normals(130, 0) = -0.147621: Normals(130, 1) = -0.716567: Normals(130, 2) = 0.681718
    Normals(131, 0) = -0.425325: Normals(131, 1) = -0.688191: Normals(131, 2) = 0.587785
    Normals(132, 0) = -0.16246:  Normals(132, 1) = -0.262866: Normals(132, 2) = 0.951056
    Normals(133, 0) = 0.442863:  Normals(133, 1) = -0.238856: Normals(133, 2) = 0.864188
    Normals(134, 0) = 0.16246:   Normals(134, 1) = -0.262866: Normals(134, 2) = 0.951056
    Normals(135, 0) = 0.309017:  Normals(135, 1) = -0.5:      Normals(135, 2) = 0.809017
    Normals(136, 0) = 0.147621:  Normals(136, 1) = -0.716567: Normals(136, 2) = 0.681718
    Normals(137, 0) = 0#:        Normals(137, 1) = -0.525731: Normals(137, 2) = 0.850651
    Normals(138, 0) = 0.425325:  Normals(138, 1) = -0.688191: Normals(138, 2) = 0.587785
    Normals(139, 0) = 0.587785:  Normals(139, 1) = -0.425325: Normals(139, 2) = 0.688191
    Normals(140, 0) = 0.688191:  Normals(140, 1) = -0.587785: Normals(140, 2) = 0.425325
    Normals(141, 0) = -0.955423: Normals(141, 1) = 0.295242:  Normals(141, 2) = 0#
    Normals(142, 0) = -0.951056: Normals(142, 1) = 0.16246:   Normals(142, 2) = 0.262866
    Normals(143, 0) = -1#:       Normals(143, 1) = 0#:        Normals(143, 2) = 0#
    Normals(144, 0) = -0.850651: Normals(144, 1) = 0#:        Normals(144, 2) = 0.525731
    Normals(145, 0) = -0.955423: Normals(145, 1) = -0.295242: Normals(145, 2) = 0#
    Normals(146, 0) = -0.951056: Normals(146, 1) = -0.16246:  Normals(146, 2) = 0.262866
    Normals(147, 0) = -0.864188: Normals(147, 1) = 0.442863:  Normals(147, 2) = -0.238856
    Normals(148, 0) = -0.951056: Normals(148, 1) = 0.16246:   Normals(148, 2) = -0.262866
    Normals(149, 0) = -0.809017: Normals(149, 1) = 0.309017:  Normals(149, 2) = -0.5
    Normals(150, 0) = -0.864188: Normals(150, 1) = -0.442863: Normals(150, 2) = -0.238856
    Normals(151, 0) = -0.951056: Normals(151, 1) = -0.16246:  Normals(151, 2) = -0.262866
    Normals(152, 0) = -0.809017: Normals(152, 1) = -0.309017: Normals(152, 2) = -0.5
    Normals(153, 0) = -0.681718: Normals(153, 1) = 0.147621:  Normals(153, 2) = -0.716567
    Normals(154, 0) = -0.681718: Normals(154, 1) = -0.147621: Normals(154, 2) = -0.716567
    Normals(155, 0) = -0.850651: Normals(155, 1) = 0#:        Normals(155, 2) = -0.525731
    Normals(156, 0) = -0.688191: Normals(156, 1) = 0.587785:  Normals(156, 2) = -0.425325
    Normals(157, 0) = -0.587785: Normals(157, 1) = 0.425325:  Normals(157, 2) = -0.688191
    Normals(158, 0) = -0.425325: Normals(158, 1) = 0.688191:  Normals(158, 2) = -0.587785
    Normals(159, 0) = -0.425325: Normals(159, 1) = -0.688191: Normals(159, 2) = -0.587785
    Normals(160, 0) = -0.587785: Normals(160, 1) = -0.425325: Normals(160, 2) = -0.688191
    Normals(161, 0) = -0.688191: Normals(161, 1) = -0.587785: Normals(161, 2) = -0.425325
End Sub
