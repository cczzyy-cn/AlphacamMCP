Imports System
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace AlphaCAMRouter
	' Token: 0x02000021 RID: 33
	<TypeIdentifier()>
	<Guid("A6185F82-45D5-11D2-9868-00104B4AF281")>
	<CompilerGenerated()>
	<ComImport()>
	Public Interface IElement
		' Token: 0x06000117 RID: 279
		Sub _VtblGap1_1()

		' Token: 0x1700004F RID: 79
		' (get) Token: 0x06000118 RID: 280
		' (set) Token: 0x06000119 RID: 281
		Property CW As Boolean

		' Token: 0x0600011A RID: 282
		Sub _VtblGap2_1()

		' Token: 0x17000050 RID: 80
		' (get) Token: 0x0600011B RID: 283
		ReadOnly Property IsLine As Boolean

		' Token: 0x0600011C RID: 284
		Sub _VtblGap3_1()

		' Token: 0x17000051 RID: 81
		' (get) Token: 0x0600011D RID: 285
		' (set) Token: 0x0600011E RID: 286
		Property StartXL As Double

		' Token: 0x17000052 RID: 82
		' (get) Token: 0x0600011F RID: 287
		' (set) Token: 0x06000120 RID: 288
		Property StartYL As Double

		' Token: 0x17000053 RID: 83
		' (get) Token: 0x06000121 RID: 289
		' (set) Token: 0x06000122 RID: 290
		Property EndXL As Double

		' Token: 0x17000054 RID: 84
		' (get) Token: 0x06000123 RID: 291
		' (set) Token: 0x06000124 RID: 292
		Property EndYL As Double

		' Token: 0x06000125 RID: 293
		Sub _VtblGap4_12()

		' Token: 0x17000055 RID: 85
		' (get) Token: 0x06000126 RID: 294
		' (set) Token: 0x06000127 RID: 295
		Property CenterXL As Double

		' Token: 0x17000056 RID: 86
		' (get) Token: 0x06000128 RID: 296
		' (set) Token: 0x06000129 RID: 297
		Property CenterYL As Double

		' Token: 0x0600012A RID: 298
		Sub _VtblGap5_18()

		' Token: 0x17000057 RID: 87
		' (get) Token: 0x0600012B RID: 299
		ReadOnly Property IsRapid As Boolean
	End Interface
End Namespace
