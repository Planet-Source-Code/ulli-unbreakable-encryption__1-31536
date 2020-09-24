VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEncDec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright © 2002 UMGEDV GmbH
'
' Encryption is a sensitive area of programming and there are people out there who
' sue for the craziest reasons; so I have to make the following statements and you,
' by using this code, do agree to abide by them.
'
' ******************************************************************************
' *  This software is provided "AS IS" and any express or implied warranties,  *
' *  including, but not limited to, the implied warranties of merchantability  *
' *  or fitness for a particular purpose, are disclaimed. In no event shall    *
' *  the author be liable for any direct, indirect, incidental, special,       *
' *  exemplary, or consequential damages (including, but not limited to,       *
' *  procurement of substitute goods or services; loss of use, data, or        *
' *  profits; or business interruption) however caused and on any theory of    *
' *  liability, whether in contract, strict liability, or tort (including      *
' *  negligence or otherwise) arising in any way out of the use of, or the     *
' *  impossibility to use, this software, even if advised of the possibility   *
' *  of such damage.                                                           *
' ******************************************************************************

' The following restrictions apply to this program code:
'
' 1 You may use this code in non-commercial applications for your own personal use.
'   If you want to use this code, or the ideas and techniques upon which this code
'   is based, in a commercial application, prior permission in writing by the author
'   is required. You can contact me via email at umgedv@aol.com
'
' 2 Any restrictions imposed by your government or other authorities regarding the
'   use of "strong encryption algorithms" must be obeyed.
'
' 3 No commitment (express or implied) is made whatsoever regarding the encryption
'   strength or security. When you use this code, privat or commercial, you do so
'   completely at your own risk.
'
' 4 These notes and restrictions must not be removed from this file.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Class...
'
' ...to encrypt and decrypt a clear text with variable character length encoding using
' symetrical en/decryption - both use the same key
'
' and...
'
' ...to encapsulate a set of MD5 Message Digest functions (based on code by
' Robert Hubley, who in turn apparently used the original C-code published by
' RSA Data Security Inc.)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' There is no relationship between byte boundaries and encrypted characters, ie an
' encrypted character may theoretically be anything from 1 to 255 bits long; however
' this depends largely on the 'quality' of the key. There are no markers or the like
' to delimit the variable length characters, the code is self-delimiting (when you know
' what to look for *g*). I am not a cryptoanalyst, but I have a hunch that the variable
' character length makes this code rather difficult if not impossible to break.
'
' One approach may be to scan the encryption for repeating bit-groups and guess the
' meaning of short groups occuring more frequently than others (like "the", "a", "an"
' and so forth), or for patterns occuring with a language-specific frequency of vowels
' and consonants. To forestall that, the encrypting engine uses an initial value (also
' derived from the key) and XORs every character with that before encrypting, keeping
' the result of the XOR for the next character. The decrypting engine of course reverses
' that process.
'
' The MD5 message digest can be used to sign a message. An advantage of these addition-
' ally scrambled bytes lies in the fact, that the encpyption engine uses a value derived
' from the previously encrypted characters to scramble the current byte, and since the
' signature is an unpredictable character sequence, so will be the scrambled clear text.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' The MD5 algorithm produces a 128 bit digital fingerprint (signature) from a
' dataset of arbitrary length. For details see RFC 1321 (summarized below). This
' implementation is derived from the RSA Data Security, Inc. MD5 Message-Digest
' algorithm reference implementation (originally written in C).
'
' NOTES:
'     Network Working Group                                             R. Rivest
'     Request for Comments: 1321              MIT Laboratory for Computer Science
'                                                      and RSA Data Security, Inc.
'                                                                      April 1992
'
'                          The MD5 Message-Digest Algorithm
'                          ================================
'     Summary
'     -------
'
'        This document describes the MD5 message-digest algorithm. The
'        algorithm takes as input a message of arbitrary length and produces
'        as output a 128-bit "fingerprint" or "message digest" of the input.
'        It is conjectured that it is computationally infeasible to produce
'        two messages having the same message digest, or to produce any
'        message having a given prespecified target message digest. The MD5
'        algorithm is intended for digital signature applications, where a
'        large file must be "compressed" in a secure manner before being
'        encrypted with a private (secret) key under a public-key cryptosystem
'        such as RSA.
'
'        The MD5 algorithm is designed to be quite fast on 32-bit machines. In
'        addition, the MD5 algorithm does not require any large substitution
'        tables; the algorithm can be coded quite compactly.
'
'        The MD5 algorithm is an extension of the MD4 message-digest algorithm.
'        MD5 is slightly slower than MD4, but is more "conservative" in
'        design. MD5 was designed because it was felt that MD4 was perhaps
'        being adopted for use more quickly than justified by the existing
'        critical review; because MD4 was designed to be exceptionally fast,
'        it is "at the edge" in terms of risking successful cryptanalytic
'        attack. MD5 backs off a bit, giving up a little in speed for a much
'        greater likelihood of ultimate security. It incorporates some
'        suggestions made by various reviewers, and contains additional
'        optimizations. The MD5 algorithm is being placed in the public domain
'        for review and possible adoption as a standard.
'
'        RFC Author:
'        Ronald L.Rivest
'        Massachusetts Institute of Technology
'        Laboratory for Computer Science
'        NE43-324545 Technology Square
'        Cambridge, MA 02139-1986
'        Phone: (617) 253-5880
'        EMail: Rivest@theory.lcs.mit.edu
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Dev History:
' ------------
'
' Feb 06 2002 UMG
'
' 1 The weight formula would assign a zero weight to a node sometimes rendering the
'   tree unusable. Changed to " Do ... Loop Until .Weight ". Also modified the
'   algorithm and constants to avoid repeated patterns of the weight and make it more
'   chaotic, resulting in a better unbalance of the tree.
'
' 2 Shifting right of hash always left the MSB zero. This resulted in a repeating
'   cycle when encrypting long sequences of equal bytes.
'   Added " + Rnd * 128 " modify hash (including MSB) randomly. Note that the
'   Randomizer is in a defined but unpredictable state after processing of the key.
'
' 3 Now Encrypting and Decrypting use a chunk size and append when the chunk size
'   is reached. This speeded up both processes significantly.
'
' 4 Slack bits at the end of an encrypted text are undecryptable now.
'
' 5 New Property BytesPerSecond (using GetTickCount() with an accuracy of 1 msec).
'
' 6 Moved mask creation to Initialize, this also speeded up decrytion.
'
' 7 Added ShowNodeWithChildren tree display.
'
' 8 Added MD5 message digest functions.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Jan 20 2002 UMG
'
'   Prototype
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
DefLng A-Z 'we're 32 bit!

#Const Debugging = True 'Set to False for final compilation

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'for variable length encryption

Private Type TreeNode   'the tree nodes
    Weight              As Long
    Parent              As Long
    Child1              As Long
    Child2              As Long
End Type

Private Nodes()         As TreeNode     'the tree
Attribute Nodes.VB_VarDescription = "The Tree"
Private Codes()         As String       'the variable lenght codes
Attribute Codes.VB_VarDescription = "The variable lenght Codes"
Private Masks(0 To 7)   As Long
Attribute Masks.VB_VarDescription = "Bit masks used during decrypting"
Private Hash            As Long
Attribute Hash.VB_VarDescription = "Password and text digest"
Private IxSlack         As Long         'index pointing to code to be used as slack
Attribute IxSlack.VB_VarDescription = "Index pointing to a suitable caditate to fill slack bits"
Private MinLen          As Long         'tree properties
Attribute MinLen.VB_VarDescription = "local property"
Private MaxLen          As Long
Attribute MaxLen.VB_VarDescription = "local property"
Private Speed           As Long         'conversion speed
Attribute Speed.VB_VarDescription = "Conversion speed"

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private StartTick       As Long
Attribute StartTick.VB_VarDescription = "From GetTickCount"

Private Const MaxLong   As Long = 2 ^ 31 - 1
Attribute MaxLong.VB_VarDescription = "2 ^ 31 - 1"
Private Const ChunkSize As Long = 255
Attribute ChunkSize.VB_VarDescription = "Chunksize is used during encryption and decryption"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'for MD5 message digest

Private Const IniPart1  As Double = 1732584193#
Private Const IniPart2  As Double = 4023233417#
Private Const IniPart3  As Double = 2562383102#
Private Const IniPart4  As Double = 271733878#

'Pass 1
Private Const S101      As Long = 7
Private Const S102      As Long = 12
Private Const S103      As Long = 17
Private Const S104      As Long = 22
Private Const A101      As Long = -680876936
Private Const A102      As Long = -389564586
Private Const A103      As Long = 606105819
Private Const A104      As Long = -1044525330
Private Const A105      As Long = -176418897
Private Const A106      As Long = 1200080426
Private Const A107      As Long = -1473231341
Private Const A108      As Long = -45705983
Private Const A109      As Long = 1770035416
Private Const A110      As Long = -1958414417
Private Const A111      As Long = -42063
Private Const A112      As Long = -1990404162
Private Const A113      As Long = 1804603682
Private Const A114      As Long = -40341101
Private Const A115      As Long = -1502002290
Private Const A116      As Long = 1236535329

'Pass 2
Private Const S201      As Long = 5
Private Const S202      As Long = 9
Private Const S203      As Long = 14
Private Const S204      As Long = 20
Private Const A201      As Long = -165796510
Private Const A202      As Long = -1069501632
Private Const A203      As Long = 643717713
Private Const A204      As Long = -373897302
Private Const A205      As Long = -701558691
Private Const A206      As Long = 38016083
Private Const A207      As Long = -660478335
Private Const A208      As Long = -405537848
Private Const A209      As Long = 568446438
Private Const A210      As Long = -1019803690
Private Const A211      As Long = -187363961
Private Const A212      As Long = 1163531501
Private Const A213      As Long = -1444681467
Private Const A214      As Long = -51403784
Private Const A215      As Long = 1735328473
Private Const A216      As Long = -1926607734

'Pass 3
Private Const S301      As Long = 4
Private Const S302      As Long = 11
Private Const S303      As Long = 16
Private Const S304      As Long = 23
Private Const A301      As Long = -378558
Private Const A302      As Long = -2022574463
Private Const A303      As Long = 1839030562
Private Const A304      As Long = -35309556
Private Const A305      As Long = -1530992060
Private Const A306      As Long = 1272893353
Private Const A307      As Long = -155497632
Private Const A308      As Long = -1094730640
Private Const A309      As Long = 681279174
Private Const A310      As Long = -358537222
Private Const A311      As Long = -722521979
Private Const A312      As Long = 76029189
Private Const A313      As Long = -640364487
Private Const A314      As Long = -421815835
Private Const A315      As Long = 530742520
Private Const A316      As Long = -995338651

'Pass 4
Private Const S401      As Long = 6
Private Const S402      As Long = 10
Private Const S403      As Long = 15
Private Const S404      As Long = 21
Private Const A401      As Long = -198630844
Private Const A402      As Long = 1126891415
Private Const A403      As Long = -1416354905
Private Const A404      As Long = -57434055
Private Const A405      As Long = 1700485571
Private Const A406      As Long = -1894986606
Private Const A407      As Long = -1051523
Private Const A408      As Long = -2054922799
Private Const A409      As Long = 1873313359
Private Const A410      As Long = -30611744
Private Const A411      As Long = -1560198380
Private Const A412      As Long = 1309151649
Private Const A413      As Long = -145523070
Private Const A414      As Long = -1120210379
Private Const A415      As Long = 718787259
Private Const A416      As Long = -343485551

Private Const Top16Bits As Long = &HFFFF0000
Private Const Bot16Bits As Long = &HFFFF&
Private Const Bot8Bits  As Long = &HFF&
Private Const Bit16Only As Long = &H10000

Private Const BlockSize As Long = 64

'variables, types, and enums
Private Type Segs
    P(1 To 4)           As Long
End Type
Private Digest          As Segs
Public Enum SignatureLength
    [Short] = 4
    [Medium] = 16
    [Long] = 32
End Enum

'local properties
Private myMsgDigest4    As String
Private myMsgDigest16   As String
Private myMsgDigest32   As String

Private Function BuildLeaves(Key As String) As Long

  'this builds all nodes for a tree, the weight assigned to the leaves however depends on
  'the key rather than on the character frequency (which would be used in a Huffman Tree)

  Dim IxLeaf, IxKey, TotalWeight, Rand

#If Debugging Then
    fTest.tvwTree.Nodes.Clear
    fTest.lstCodes.Clear
#End If
    ReDim Nodes(0 To 510)

    If Len(Key) Then
        For IxLeaf = LBound(Nodes) To UBound(Nodes)
            If IxKey = Len(Key) Then
                IxKey = 0
            End If
            IxKey = IxKey + 1 'cycles through key
            With Nodes(IxLeaf)
                .Parent = 0 'no parents or children yet
                .Child1 = 0
                .Child2 = 0
                If IxLeaf < 256 Then 'this is a leaf for sure

                    'this rather complicated algorithm assigns a weight to a leaf, this weight should
                    'be widely spread to make the tree unbalanced, and chaotic to react to input
                    'changes with 'unpredictable' output changes

                    Rand = Int(Rnd(-Asc(Mid$(Key, IxKey, 1)) - IxLeaf Xor IxKey) * 2674317)
                    Do
                        .Weight = (Rand + IxLeaf) Mod (Rnd ^ 11 * 473 * Asc(Mid$(Key, Int(Rnd * Len(Key) + 1), 1)) + 1)
                    Loop Until .Weight

                    'we need the total weight of all leaves to recognize the root later on
                    TotalWeight = TotalWeight + .Weight

                  Else 'this will be a parent node'NOT IXLEAF...
                    .Weight = 0 'will later have the combined weight of her children
                End If
            End With 'NODES(IXLEAF)
        Next IxLeaf
        BuildLeaves = TotalWeight 'return total weight of all leaves
      Else 'LEN(KEY) = FALSE
        Err.Raise 1001, "cEncDec", "The key is missing."
    End If

    'create initial hash
    Hash = (TotalWeight Xor TotalWeight \ 256 Xor TotalWeight \ 256 \ 256 Xor TotalWeight \ 256 \ 256 \ 256) And Bot8Bits
    'init randomizer for hash
    Rnd -TotalWeight * Hash

End Function

Private Function BuildTree(ExpectedRootWeight As Long) As Long

  'this creates a tree in Huffman fashion, however the node-weights do not depend on
  'character frequency but on the key (see BuildLeaves)

  Dim ChildCount, IxNode, IxParent, IxFirstChild, IxSecondChild, LowestWeight

    IxParent = 255 'leaves are in 0 .. 255, parents, grandparents (and finally root) will be in 256 .. 512
    Do
        IxParent = IxParent + 1
        With Nodes(IxParent)
            For ChildCount = 1 To 2 'find two children for the next parent
                IxFirstChild = IxSecondChild '2nd index is free again on 2nd time Pass
                LowestWeight = MaxLong
                For IxNode = 0 To IxParent - 1 'find parentless living Nodes with the lowest Weight (leaf or intermediate)
                    If Nodes(IxNode).Parent = 0 And Nodes(IxNode).Weight > 0 And Nodes(IxNode).Weight < LowestWeight Then
                        LowestWeight = Nodes(IxNode).Weight
                        IxSecondChild = IxNode 'remember that child's index
                    End If
                Next IxNode
                If LowestWeight < MaxLong Then 'found a suitable child - adopt it
                    Nodes(IxSecondChild).Parent = IxParent 'tell child who her parent is
                    .Weight = .Weight + LowestWeight 'add child weight to parent weight
                End If
            Next ChildCount
            .Child1 = IxFirstChild 'tell parent who her children are
            .Child2 = IxSecondChild
        End With 'NODES(IXPARENT)
    Loop Until Nodes(IxParent).Weight = ExpectedRootWeight 'bail out if the root was just created
    BuildTree = IxParent 'return root index
#If Debugging Then
    ShowNodeWithChildren -1, IxParent, vbNullString, vbNullString
#End If

End Function

Public Property Get BytesPerSecond() As Long
Attribute BytesPerSecond.VB_Description = "Public Property"

    BytesPerSecond = Speed

End Property

Private Sub Class_Initialize()

  Dim IxMask

    For IxMask = LBound(Masks) To UBound(Masks)
        Masks(IxMask) = 2 ^ IxMask 'Masks are initialized only once; speed up decryption
    Next IxMask

End Sub

Private Sub Class_Terminate()

  'erase sensitive info form memory

    Erase Nodes, Codes, Digest.P

End Sub

Private Sub CreateCodes(Root As Long)

  'this scans the tree from leaves to root and creates a table of variable length codes

  Dim IxLeaf, IxChild, IxParent, Code As String

    ReDim Codes(0 To 255)
    MinLen = MaxLong
    MaxLen = -1
    For IxLeaf = LBound(Codes) To UBound(Codes) 'all leaves
        Code = vbNullString
        If Nodes(IxLeaf).Parent Then 'not an orphan (don't think there are any in this tree but just to be sure)
            IxChild = IxLeaf
            Do 'climb up to root
                IxParent = Nodes(IxChild).Parent
                Code = IIf(IxChild = Nodes(IxParent).Child1, "0", "1") & Code
                IxChild = IxParent
            Loop Until IxParent = Root
            Codes(IxLeaf) = Code
            If Len(Code) < MinLen Then
                MinLen = Len(Code)
            End If
            If Len(Code) > MaxLen Then
                MaxLen = Len(Code)
            End If
            If Len(Code) > 7 Then 'if the code is longer than 7 bits then
                IxSlack = IxLeaf  'that's a good candidate for the slack bits
            End If
        End If
    Next IxLeaf

End Sub

Private Function CreateTree(Key As String) As Long

    CreateTree = BuildTree(BuildLeaves(Key))

End Function

Public Function Decrypt(Data As String, Key As String) As String
Attribute Decrypt.VB_Description = "This reads the input bit stream and climbs down the tree until a leaf is found; that leaf's index (after digest) is the Ascii value of clear character to return"

  'this reads the input bit stream and climbs down the tree until a leaf is found;
  'that leaf's index (after hash) is the Ascii value of clear character to return

  Dim Root, IxParent, IxChild, Pointer, IxMask, Coding, Ascii, Chunk As String

    StartTick = GetTickCount - 1 ' (-1) to prevent zero division
    Root = CreateTree(Key)

    'decrypt data
    IxParent = Root 'start at root
    For Pointer = 1 To Len(Data)
        Coding = Asc(Mid$(Data, Pointer, 1))
        For IxMask = UBound(Masks) To LBound(Masks) Step -1 'climb down the tree
            If Coding And Masks(IxMask) Then
                IxChild = Nodes(IxParent).Child2
              Else 'NOT CODING...
                IxChild = Nodes(IxParent).Child1
            End If
            If Nodes(IxChild).Child1 = 0 And Nodes(IxChild).Child2 = 0 Then 'that's a leaf
                Ascii = IxChild Xor Hash
                Chunk = Chunk & Chr$(Ascii) 'append clear char
                Hash = (Ascii Xor Hash) \ 2 + Rnd * 128 'update hash
                IxParent = Root 'go back to root
                If Len(Chunk) = ChunkSize Then
                    Decrypt = Decrypt & Chunk 'append chunk
                    Chunk = vbNullString
                End If
              Else 'not a leaf'NOT NODES(IXCHILD).CHILD1...
                IxParent = IxChild 'climb down further
            End If
    Next IxMask, Pointer
    Decrypt = Decrypt & Chunk

    Speed = Len(Decrypt) * 1000 / (GetTickCount - StartTick)
    Class_Terminate

End Function

Private Sub DigestBlock(Data As String)

  'MD5 mangle, requires a string of length 64

  Dim IxBlock, Packed(0 To 15)
  Dim a, b, c, d

    'pack 4 bytes into 1 long
    For IxBlock = 1 To BlockSize Step 4
        Packed(IxBlock \ 4) = RotateLeft(RotateLeft(RotateLeft(Asc(Mid$(Data, IxBlock + 3, 1)), 8) Or Asc(Mid$(Data, IxBlock + 2, 1)), 8) Or Asc(Mid$(Data, IxBlock + 1, 1)), 8) Or Asc(Mid$(Data, IxBlock, 1))
    Next IxBlock
    
    With Digest
        'Pass 1 digest
        a = Pass1(.P(1), .P(2), .P(3), .P(4), Packed(0), S101, A101)
        d = Pass1(.P(4), a, .P(2), .P(3), Packed(1), S102, A102)
        c = Pass1(.P(3), d, a, .P(2), Packed(2), S103, A103)
        b = Pass1(.P(2), c, d, a, Packed(3), S104, A104)

        a = Pass1(a, b, c, d, Packed(4), S101, A105)
        d = Pass1(d, a, b, c, Packed(5), S102, A106)
        c = Pass1(c, d, a, b, Packed(6), S103, A107)
        b = Pass1(b, c, d, a, Packed(7), S104, A108)

        a = Pass1(a, b, c, d, Packed(8), S101, A109)
        d = Pass1(d, a, b, c, Packed(9), S102, A110)
        c = Pass1(c, d, a, b, Packed(10), S103, A111)
        b = Pass1(b, c, d, a, Packed(11), S104, A112)

        a = Pass1(a, b, c, d, Packed(12), S101, A113)
        d = Pass1(d, a, b, c, Packed(13), S102, A114)
        c = Pass1(c, d, a, b, Packed(14), S103, A115)
        b = Pass1(b, c, d, a, Packed(15), S104, A116)

        'Pass 2 digest
        a = Pass2(a, b, c, d, Packed(1), S201, A201)
        d = Pass2(d, a, b, c, Packed(6), S202, A202)
        c = Pass2(c, d, a, b, Packed(11), S203, A203)
        b = Pass2(b, c, d, a, Packed(0), S204, A204)

        a = Pass2(a, b, c, d, Packed(5), S201, A205)
        d = Pass2(d, a, b, c, Packed(10), S202, A206)
        c = Pass2(c, d, a, b, Packed(15), S203, A207)
        b = Pass2(b, c, d, a, Packed(4), S204, A208)

        a = Pass2(a, b, c, d, Packed(9), S201, A209)
        d = Pass2(d, a, b, c, Packed(14), S202, A210)
        c = Pass2(c, d, a, b, Packed(3), S203, A211)
        b = Pass2(b, c, d, a, Packed(8), S204, A212)

        a = Pass2(a, b, c, d, Packed(13), S201, A213)
        d = Pass2(d, a, b, c, Packed(2), S202, A214)
        c = Pass2(c, d, a, b, Packed(7), S203, A215)
        b = Pass2(b, c, d, a, Packed(12), S204, A216)

        'Pass 3 digest
        a = Pass3(a, b, c, d, Packed(5), S301, A301)
        d = Pass3(d, a, b, c, Packed(8), S302, A302)
        c = Pass3(c, d, a, b, Packed(11), S303, A303)
        b = Pass3(b, c, d, a, Packed(14), S304, A304)

        a = Pass3(a, b, c, d, Packed(1), S301, A305)
        d = Pass3(d, a, b, c, Packed(4), S302, A306)
        c = Pass3(c, d, a, b, Packed(7), S303, A307)
        b = Pass3(b, c, d, a, Packed(10), S304, A308)

        a = Pass3(a, b, c, d, Packed(13), S301, A309)
        d = Pass3(d, a, b, c, Packed(0), S302, A310)
        c = Pass3(c, d, a, b, Packed(3), S303, A311)
        b = Pass3(b, c, d, a, Packed(6), S304, A312)

        a = Pass3(a, b, c, d, Packed(9), S301, A313)
        d = Pass3(d, a, b, c, Packed(12), S302, A314)
        c = Pass3(c, d, a, b, Packed(15), S303, A315)
        b = Pass3(b, c, d, a, Packed(2), S304, A316)

        'Pass 4 digest
        a = Pass4(a, b, c, d, Packed(0), S401, A401)
        d = Pass4(d, a, b, c, Packed(7), S402, A402)
        c = Pass4(c, d, a, b, Packed(14), S403, A403)
        b = Pass4(b, c, d, a, Packed(5), S404, A404)

        a = Pass4(a, b, c, d, Packed(12), S401, A405)
        d = Pass4(d, a, b, c, Packed(3), S402, A406)
        c = Pass4(c, d, a, b, Packed(10), S403, A407)
        b = Pass4(b, c, d, a, Packed(1), S404, A408)

        a = Pass4(a, b, c, d, Packed(8), S401, A409)
        d = Pass4(d, a, b, c, Packed(15), S402, A410)
        c = Pass4(c, d, a, b, Packed(6), S403, A411)
        b = Pass4(b, c, d, a, Packed(13), S404, A412)

        a = Pass4(a, b, c, d, Packed(4), S401, A413)
        d = Pass4(d, a, b, c, Packed(11), S402, A414)
        c = Pass4(c, d, a, b, Packed(2), S403, A415)
        b = Pass4(b, c, d, a, Packed(9), S404, A416)

        .P(1) = UnsignedAdd(.P(1), a, 0, 0)
        .P(2) = UnsignedAdd(.P(2), b, 0, 0)
        .P(3) = UnsignedAdd(.P(3), c, 0, 0)
        .P(4) = UnsignedAdd(.P(4), d, 0, 0)
    End With 'DIGEST

End Sub

Public Function Encrypt(Data As String, Key As String) As String
Attribute Encrypt.VB_Description = "Here the created codes are used to build the output bit stream"

  'here the created codes are used to build the output bit stream

  Dim Pointer, IxCode, Bits, Code As String, Chunk As String

    StartTick = GetTickCount - 1 ' (-1) to prevent zero division
    CreateCodes CreateTree(Key)

    'encrypt data
    Bits = 1 'marker bit
    For Pointer = 1 To Len(Data)
        IxCode = Asc(Mid$(Data, Pointer, 1)) Xor Hash
        Code = Codes(IxCode)
        Hash = IxCode \ 2 + Rnd * 128
#If Debugging Then
        fTest.lstCodes.AddItem Mid$(Data, Pointer, 1) & " = " & Code & " (" & Len(Code) & ")"
#End If
        For IxCode = 1 To Len(Code)
            Bits = Bits * 2
            If Mid$(Code, IxCode, 1) = "1" Then
                Bits = Bits + 1
            End If
            If Bits > Bot8Bits Then
                Chunk = Chunk & Chr$(Bits And Bot8Bits) 'append 8 encrypted bits
                Bits = 1 'marker bit
            End If
            If Len(Chunk) = ChunkSize Then
                Encrypt = Encrypt & Chunk 'append chunk
                Chunk = vbNullString
            End If
    Next IxCode, Pointer
    If Bits > 1 Then 'append remaining bits and slack bits
#If Debugging Then
        Code = ""
#End If
        Do
            Bits = Bits * 2
            If Left$(Codes(IxSlack), 1) = "1" Then
                Bits = Bits + 1
#If Debugging Then
                Code = Code & "1"
              Else 'NOT LEFT$(CODES(IXSLACK),...
                Code = Code & "0"
#End If
            End If
            Codes(IxSlack) = Mid$(Codes(IxSlack), 2) 'don't need code no more after this, can destroy
        Loop Until Bits > Bot8Bits
#If Debugging Then
        fTest.lstCodes.AddItem "[slack] = " & Code & " (" & Len(Code) & ")"
#End If
        Chunk = Chunk & Chr$(Bits And Bot8Bits)
    End If
    Encrypt = Encrypt & Chunk
    Speed = Len(Data) * 1000 / (GetTickCount - StartTick)
    Class_Terminate

End Function

Public Property Get LastSignature(Length As SignatureLength) As String

    LastSignature = Switch(Length = [Short], myMsgDigest4, Length = [Medium], myMsgDigest16, Length = [Long], myMsgDigest32, True, "")
    If LastSignature = "" Then
        Err.Raise 1003, , "Signature Lenght " & Length & " is not available."
    End If

End Property

Public Property Get MaxCodeLength() As Long
Attribute MaxCodeLength.VB_Description = "Tree property: longest path from root to leaf"

  'returns path length from root to farthest leaf

    MaxCodeLength = MaxLen

End Property

Public Property Get MinCodeLength() As Long
Attribute MinCodeLength.VB_Description = "Tree property: shortest path from root to leaf"

  'returns path length from root to nearest leaf

    MinCodeLength = MinLen

End Property

Private Function Pass1(a As Long, b As Long, c As Long, d As Long, x As Long, s As Long, Ac As Long) As Long

    Pass1 = UnsignedAdd(RotateLeft(UnsignedAdd(a, (b And c) Or (d And Not (b)), x, Ac), s), b, 0, 0)

End Function

Private Function Pass2(a As Long, b As Long, c As Long, d As Long, x As Long, s As Long, Ac As Long) As Long

    Pass2 = UnsignedAdd(RotateLeft(UnsignedAdd(a, (b And d) Or (c And Not (d)), x, Ac), s), b, 0, 0)

End Function

Private Function Pass3(a As Long, b As Long, c As Long, d As Long, x As Long, s As Long, Ac As Long) As Long

    Pass3 = UnsignedAdd(RotateLeft(UnsignedAdd(a, b Xor c Xor d, x, Ac), s), b, 0, 0)

End Function

Private Function Pass4(a As Long, b As Long, c As Long, d As Long, x As Long, s As Long, Ac As Long) As Long

    Pass4 = UnsignedAdd(RotateLeft(UnsignedAdd(a, c Xor (b Or Not (d)), x, Ac), s), b, 0, 0)

End Function

Private Function RotateLeft(ByVal Value As Long, ByVal NumPlaces As Long) As Long

  'rotate a long numplaces to the left

  Dim CntPlaces, SpillOver

    For CntPlaces = 1 To NumPlaces Mod 32
        SpillOver = Value And &HC0000000
        Value = (Value And &H3FFFFFFF) * 2 Or ((SpillOver < 0) And 1) Or (CBool(SpillOver And &H40000000) And &H80000000)
    Next CntPlaces
    RotateLeft = Value

End Function

#If Debugging Then
Private Sub ShowNodeWithChildren(IxParent As Long, IxNode As Long, ChildNum As String, Coding As String)

  'show a graphic representation of the tree

  Dim PKey As String, CKey As String

    PKey = "K" & IxParent
    CKey = "K" & IxNode
    With fTest.tvwTree.Nodes
        If IxParent < 0 Then
            .Add(, , CKey, IxNode, "Root").Expanded = False
          Else 'NOT IXPARENT...
            If Nodes(IxNode).Child1 Then
                .Add(PKey, tvwChild, CKey, IxNode & " (C" & ChildNum & " of " & IxParent & ")", "Node").Expanded = True
              Else 'NODES(IXNODE).CHILD1 = FALSE
                .Add(PKey, tvwChild, CKey, IxNode & " (C" & ChildNum & " of " & IxParent & ")" & " = " & Coding, "Leaf").Expanded = True
            End If
        End If
    End With 'FTEST.TVWTREE.NODES
    With Nodes(IxNode)
        If .Child1 Then
            ShowNodeWithChildren IxNode, .Child1, "1", Coding & "0"
        End If
        If .Child2 Then
            ShowNodeWithChildren IxNode, .Child2, "2", Coding & "1"
        End If
    End With 'NODES(IXNODE)

End Sub
#End If

Public Function Signature(Reset As Boolean, Data As String, LenSignature As SignatureLength) As String
Attribute Signature.VB_Description = "Digest a string and output the result as a string or a hex string"

  'digest a string and output the result as a string
  'this can be called several times in a row, for very long texts which have to be split
  'first call for must reset, subsequent calls must not

  Dim ip, Padding As String

    With Digest
        If Reset Then
            'initialize the algorithm
            .P(1) = ToUnsignedLong(IniPart1)
            .P(2) = ToUnsignedLong(IniPart2)
            .P(3) = ToUnsignedLong(IniPart3)
            .P(4) = ToUnsignedLong(IniPart4)
        End If

        'create pseudo random padding
        Padding = Mid$(Data, (Len(Data) + 1) \ 2, 1)
        Do
            Padding = Mid$(Data, Asc(Padding) Mod Len(Data) + 1, 1) & Padding
        Loop Until Len(Padding) = BlockSize - 1

        'digest the stuff in chunks
        For ip = 1 To Len(Data) Step BlockSize
            DigestBlock Mid$(Data & Padding, ip, BlockSize)
        Next ip

        'make three different signatures:
        '4 bytes
        ip = .P(1) Xor .P(2) Xor .P(3) Xor .P(4)
        myMsgDigest4 = Chr$(ip And Bot8Bits) & Chr$(ip \ 256 And Bot8Bits) & Chr$(ip \ 256 \ 256 And Bot8Bits) & Chr$(ip \ 256 \ 256 \ 256 And Bot8Bits)
    
        '16 bytes
        myMsgDigest16 = ""
        For ip = 0 To 3
            myMsgDigest16 = myMsgDigest16 & Chr$(.P(1) \ 256 ^ ip And Bot8Bits) & Chr$(.P(2) \ 256 ^ ip And Bot8Bits) & Chr$(.P(3) \ 256 ^ ip And Bot8Bits) & Chr$(.P(4) \ 256 ^ ip And Bot8Bits)
        Next ip

        '32 bytes
        myMsgDigest32 = LCase$(Right$("0000000" & Hex$(.P(1)), 8) & Right$("0000000" & Hex$(.P(2)), 8) & Right$("0000000" & Hex$(.P(3)), 8) & Right$("0000000" & Hex$(.P(4)), 8))
    End With 'DIGEST
    Select Case LenSignature 'which should I return
      Case [Short]
        Signature = myMsgDigest4
      Case [Medium]
        Signature = myMsgDigest16
      Case [Long]
        Signature = myMsgDigest32
      Case Else
        Signature = ""
        Err.Raise 1003, , "Signature Lenght " & LenSignature & " is not available."
    End Select

End Function

Private Function ToUnsignedLong(Value As Double) As Long

  'convert a double to an unsigned long

  Const Over As Double = 2 ^ 32

    If Value <= MaxLong Then
        ToUnsignedLong = Value
      Else 'NOT VALUE...
        ToUnsignedLong = Value - Over
    End If

End Function

Private Function UnsignedAdd(Summand1 As Long, Summand2 As Long, Summand3 As Long, Summand4 As Long) As Long

  'adds four unsigned numbers together, overflows are ignored

  Dim Low, High

    Low = (Summand1 And Bot16Bits) + (Summand2 And Bot16Bits) + (Summand3 And Bot16Bits) + (Summand4 And Bot16Bits)
    High = ((Summand1 And Top16Bits) \ Bit16Only + (Summand2 And Top16Bits) \ Bit16Only + (Summand3 And Top16Bits) \ Bit16Only + (Summand4 And Top16Bits) \ Bit16Only + Low \ Bit16Only) And Bot16Bits
    UnsignedAdd = ToUnsignedLong(CDbl(High) * Bit16Only + (Low And Bot16Bits))

End Function

Public Property Get Version() As String
 
    Version = "2.0.17, dated Feb 06, 2002"

End Property

':) Ulli's VB Code Formatter V2.9.4 (06.02.2002 12:44:58) 310 + 542 = 852 Lines
