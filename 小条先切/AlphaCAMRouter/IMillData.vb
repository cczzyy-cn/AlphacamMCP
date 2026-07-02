Imports System
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace AlphaCAMRouter
	' Token: 0x02000025 RID: 37
	<TypeIdentifier()>
	<CompilerGenerated()>
	<Guid("CBA3E9D6-4A22-11D2-9874-00104B4AF281")>
	<ComImport()>
	Public Interface IMillData
		' Token: 0x17000058 RID: 88
		' (get) Token: 0x06000130 RID: 304
		' (set) Token: 0x06000131 RID: 305
		Property FinalDepth As Single

		' Token: 0x17000059 RID: 89
		' (get) Token: 0x06000132 RID: 306
		' (set) Token: 0x06000133 RID: 307
		Property SafeRapidLevel As Single

		' Token: 0x1700005A RID: 90
		' (get) Token: 0x06000134 RID: 308
		' (set) Token: 0x06000135 RID: 309
		Property RapidDownTo As Single

		' Token: 0x06000136 RID: 310
		Sub _VtblGap1_10()

		' Token: 0x1700005B RID: 91
		' (get) Token: 0x06000137 RID: 311
		' (set) Token: 0x06000138 RID: 312
		Property SpindleSpeed As Single

		' Token: 0x1700005C RID: 92
		' (get) Token: 0x06000139 RID: 313
		' (set) Token: 0x0600013A RID: 314
		Property DownFeed As Single

		' Token: 0x1700005D RID: 93
		' (get) Token: 0x0600013B RID: 315
		' (set) Token: 0x0600013C RID: 316
		Property CutFeed As Single

		' Token: 0x0600013D RID: 317
		Sub _VtblGap2_25()

		' Token: 0x1700005E RID: 94
		' (get) Token: 0x0600013E RID: 318
		' (set) Token: 0x0600013F RID: 319
		Property DepthsOfCutSpecified As Boolean

		' Token: 0x06000140 RID: 320
		Sub _VtblGap3_46()

		' Token: 0x06000141 RID: 321
		<DispId(48)>
		Function RoughFinish() As <MarshalAs(UnmanagedType.[Interface])> Paths

		' Token: 0x06000142 RID: 322
		Sub _VtblGap4_3()

		' Token: 0x06000143 RID: 323
		<DispId(52)>
		Function ManualToolPath(<[In]()> StartX As Double, <[In]()> StartY As Double, <[In]()> StartZ As Double) As <MarshalAs(UnmanagedType.[Interface])> MillManualToolPath
	End Interface
End Namespace
