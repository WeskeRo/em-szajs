Dim oACL As New MFilesAPI.AccessControlList
Dim oACEK As New MFilesAPI.AccessControlEntryKey
Dim oACED As New MFilesAPI.AccessControlEntryData
Dim oIPIDLevel As New MFilesAPI.IndirectPropertyIDLevel

' Full control for department (Department Property ID = 1142)
oACED.SetAllPermissions(MFilesAPI.MFPermission.MFPermissionAllow)
oACEK.PseudoUserID = New MFilesAPI.IndirectPropertyID
oIPIDLevel.LevelType = MFilesAPI.MFIndirectPropertyIDLevelType._
MFIndirectPropertyIDLevelTypePropertyDef
oIPIDLevel.ID = 1142
oACEK.PseudoUserID.Add(-1, oIPIDLevel)
oACL.CustomComponent.AccessControlEntries.Add(oACEK, oACED)

' Read + edit access to document creator
oACED = New MFilesAPI.AccessControlEntryData
oACEK = New MFilesAPI.AccessControlEntryKey
oIPIDLevel = New MFilesAPI.IndirectPropertyIDLevel

oACED.ChangePermissionsPermission = MFilesAPI.MFPermission.MFPermissionNotSet
oACED.DeletePermission = MFilesAPI.MFPermission.MFPermissionNotSet
oACED.EditPermission = MFilesAPI.MFPermission.MFPermissionAllow
oACED.ReadPermission = MFilesAPI.MFPermission.MFPermissionAllow
oACEK.PseudoUserID = New MFilesAPI.IndirectPropertyID

oIPIDLevel.LevelType = MFilesAPI.MFIndirectPropertyIDLevelType._
MFIndirectPropertyIDLevelTypePropertyDef
oIPIDLevel.ID = 25
oACEK.PseudoUserID.Add(-1, oIPIDLevel)
oACL.CustomComponent.AccessControlEntries.Add(oACEK, oACED)

' Read access to BillR (User ID = 22)
oACED = New MFilesAPI.AccessControlEntryData
oACEK = New MFilesAPI.AccessControlEntryKey
oACED.ReadPermission = MFilesAPI.MFPermission.MFPermissionAllow
oACEK.SetUserOrGroupID(22, False)
oACL.CustomComponent.AccessControlEntries.Add(oACEK, oACED)

' Update the permissions of document id 413
Dim ObjID As New MFilesAPI.ObjID
ObjID.SetIDs(0, 413)
Dim ObjVer As MFilesAPI.ObjVer
ObjVer = oVault.ObjectOperations.GetLatestObjVer(ObjID, False, True)
oVault.ObjectOperations.ChangePermissionsToACL(ObjVer, oACL, True)





