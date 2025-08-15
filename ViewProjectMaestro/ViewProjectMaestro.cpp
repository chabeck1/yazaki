#include "pch.h"
//mksapiviewer --iplocal si viewproject
// --project="/Projects/FCA/VF_VSIM_2022/project.pj"
// --filterSubs
// --filter=attribute:Build_DJ,attribute:Build_Boot,attribute:Build_HSM,attribute:Build_HSMUP,attribute:Build_BM,attribute:Build_BU
// --fields=memberarchive,name,memberrev,cpid
// --projectRevision=D25_X1A__1_VSIM0b_bootload
using namespace System;
#include <mksapi.h>
#include <wchar.h>
#include <vcclr.h>
/* build mods
/I"C:\Program Files (x86)\Integrity\IntegrityClient11\lib\include"
/DYNAMICBASE "mksapi.lib" /LIBPATH:"C:\Program Files (x86)\Integrity\IntegrityClient11\bin"
*/
System::Int32 DBGetAddEntry(System::Data::OleDb::OleDbConnection^ db, System::String^ Table, System::String^ val)
{
    System::Data::OleDb::OleDbCommand^ cmd_get = db->CreateCommand();
    cmd_get->CommandText = L"SELECT [" + Table + "].[ID] "
                        L"FROM [" + Table + "] "
                        L"WHERE [" + Table + "].[Desc]='" + val + "'";
    //System::Diagnostics::Debug::WriteLine(cmd_get->CommandText);
    System::Int32^ v = (System::Int32 ^) cmd_get->ExecuteScalar();
    if (v == nullptr)
    {
        System::Data::OleDb::OleDbCommand^ cmd_set = db->CreateCommand();
        cmd_set->CommandText = L"INSERT INTO [" + Table + "] ([Desc]) "
                               L"VALUES ('" + val + "')";
        //System::Diagnostics::Debug::WriteLine(cmd_set->CommandText);
        System::Diagnostics::Debug::Assert (cmd_set->ExecuteNonQuery() == 1);
        v = (System::Int32^) cmd_get->ExecuteScalar();
        System::Diagnostics::Debug::Assert(v != nullptr);
    }
    return *v;
}
System::Void DBAddCheckpointMember(System::Data::OleDb::OleDbConnection^ db, 
                                    System::Int32 checkpointID,
                                    System::Int32 nameID,
                                    System::Int32 memberArchiveID,
                                    System::String^ MemberRev,
                                    System::String^ CPID)
{
    System::Data::OleDb::OleDbCommand^ cmd_set = db->CreateCommand();
    cmd_set->CommandText = L"INSERT INTO [IntegrityCheckpointMembers] "
                           L"([CheckpointID],[NameID],[MemberArchiveID],[MemberRev],[CPID]) "
                           L"VALUES (" + checkpointID + "," +
                           nameID + "," +
                           memberArchiveID + ",'" +
                           MemberRev + "','" +
                           CPID + "')";
    //System::Diagnostics::Debug::WriteLine(cmd_set->CommandText);
    System::Diagnostics::Debug::Assert(cmd_set->ExecuteNonQuery() == 1);
}
int main(array<System::String ^> ^args)
{
    System::Data::OleDb::OleDbConnection^ db = gcnew System::Data::OleDb::OleDbConnection
                                                                            //(L"Provider=Microsoft.ACE.OLEDB.12.0;"
                                                                            (L"Provider=Microsoft.Jet.OLEDB.4.0;"
                                                                             L"Data Source=C:\\"
                                                                             L"Users\\"
                                                                             L"10032877\\"
                                                                             L"Documents\\"
                                                                            //L"DJ_VSIM_2022\\"
                                                                             L"VF_VSIM_2022\\"
                                                                             L"Software Development\\"
                                                                             L"Eng\\"
                                                                             L"Test\\"
                                                                             L"Static Code Check\\"
                                                                             L"CS_00152_03 Programming Rules.mdb");

    db->Open();
    System::Int32 CheckpointID = DBGetAddEntry(db, L"IntegrityCheckpoints", args[0]);

    mksIntegrationPoint aPoint = NULL;
    mksSession          aSession = NULL;
    mksCmdRunner        run = NULL;
    System::Diagnostics::Trace::Assert(mksAPIInitialize((const char*)"PTC Runner.log") == MKS_SUCCESS);
    mksLogConfigure(MKS_LOG_WARNING, MKS_LOG_LOW);
    System::Diagnostics::Trace::Assert(mksCreateLocalAPIConnector(&aPoint, 4, 16, 0) == MKS_SUCCESS);
    System::Diagnostics::Trace::Assert(mksGetCommonSession(&aSession, aPoint) == MKS_SUCCESS);
    System::Diagnostics::Trace::Assert(mksCreateCmdRunner(&run, aSession) == MKS_SUCCESS);
    const enum mksExecuteTypeEnum ET = NO_INTERIM;

    mksCommand cmd = mksCreateCommand();
    cmd->appName = L"si";
    cmd->cmdName = L"viewproject";

    System::Diagnostics::Trace::Assert(mksOptionListAdd(cmd->optionList,
        L"project",
        L"/Projects/FCA/VF_VSIM_2022/project.pj") == MKS_SUCCESS);
    System::Diagnostics::Trace::Assert(mksOptionListAdd(cmd->optionList,
        L"filterSubs",
        NULL) == MKS_SUCCESS);
    System::Diagnostics::Trace::Assert(mksOptionListAdd(cmd->optionList,
        L"filter",
        L"attribute:Build_DJ,attribute:Build_Boot,attribute:Build_HSM,attribute:Build_HSMUP,attribute:Build_BM,attribute:Build_BU") == MKS_SUCCESS);
    System::Diagnostics::Trace::Assert(mksOptionListAdd(cmd->optionList,
        L"fields",
        L"memberarchive,name,memberrev,cpid") == MKS_SUCCESS);
    pin_ptr<const wchar_t> RevStr = PtrToStringChars(args[0]);
    System::Diagnostics::Trace::Assert(mksOptionListAdd(cmd->optionList,
        L"projectRevision",
        RevStr) == MKS_SUCCESS);
    System::Diagnostics::Trace::Assert(mksOptionListAdd(cmd->optionList,
        L"recurse",
        NULL) == MKS_SUCCESS);

    mksResponse resp = mksCmdRunnerExecCmd(run, cmd, ET);
    wchar_t cmdstr[1024];
    System::Diagnostics::Trace::Assert(mksResponseGetCompleteCommand(resp, cmdstr, 1024) == MKS_SUCCESS);
    System::Diagnostics::Debug::WriteLine(gcnew System::String(cmdstr));
    mksWorkItem it = mksResponseGetFirstWorkItem(resp);
    while (it != NULL)
    {
        wchar_t str[501];
        mksField t;
        mksItem itm;
#if 0
        System::Diagnostics::Trace::Assert(mksWorkItemGetId(it, str, 500) == MKS_SUCCESS);
        System::String^ Id = gcnew System::String(str);

        System::String^ Context = nullptr;
        if (mksWorkItemGetContext(it, str, 500) == MKS_SUCCESS)
        {
            Context = gcnew System::String(str);
        }
#endif
        System::Diagnostics::Trace::Assert(mksWorkItemGetModelType(it, str, 500) == MKS_SUCCESS);
        System::String^ ModelType = gcnew System::String(str);

        System::String^ MemberArchive = nullptr;
        t = mksWorkItemGetField(it, L"memberarchive");
        if (t != NULL)
        {
            System::Diagnostics::Trace::Assert(mksFieldGetItemValue(t, &itm) == MKS_SUCCESS);
            System::Diagnostics::Trace::Assert(mksItemGetId(itm, str, 500) == MKS_SUCCESS);
            MemberArchive = gcnew System::String(str);
        }

        System::String^ Name = nullptr;
        t = mksWorkItemGetField(it, L"name");
        if (t != NULL)
        {
            System::Diagnostics::Trace::Assert(mksFieldGetStringValue(t, str, 500) == MKS_SUCCESS);
            Name = gcnew System::String(str);
        }

        System::String^ MemberRev = nullptr;
        t = mksWorkItemGetField(it, L"memberrev");
        if (t != NULL)
        {
            System::Diagnostics::Trace::Assert(mksFieldGetItemValue(t, &itm) == MKS_SUCCESS);
            System::Diagnostics::Trace::Assert(mksItemGetId(itm, str, 500) == MKS_SUCCESS);
            MemberRev = gcnew System::String(str);
        }

        System::String^ CPID = nullptr;
        t = mksWorkItemGetField(it, L"cpid");
        if (t != NULL)
        {
            System::Diagnostics::Trace::Assert(mksFieldGetItemValue(t, &itm) == MKS_SUCCESS);
            System::Diagnostics::Trace::Assert(mksItemGetId(itm, str, 500) == MKS_SUCCESS);
            CPID = gcnew System::String(str);
        }
#if 0
        /*
        System::Console::WriteLine("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}",
                                    Id,
                                    Context,
                                    ModelType,
                                    Type,
                                    MemberArchive,
                                    Name,
                                    MemberRev,
                                    CPID);
        */
#endif
        if (ModelType == "si.Member")
        {
            System::Int32 v = DBGetAddEntry(db, L"IntegrityMembers", Name);
            System::Int32 mA = DBGetAddEntry(db, L"IntegrityArchives", MemberArchive);
            DBAddCheckpointMember(db, CheckpointID, v, mA, MemberRev, CPID);
        }
        it = mksResponseGetNextWorkItem(resp);
    }
    mksReleaseCommand(cmd);
    mksReleaseResponse(resp);
    mksReleaseCmdRunner(run);
    db->Close();
    return 0;
}
/*
Work Item :
Id = license.liz
Context = / Projects / FCA / VF_VSIM_2022 / Software Development / Eng / Bootload / CBD2000646_D00 / Misc / HexView / project.pj
Model Type = si.Member
Field :
Name = type
Data Type = wchar_t*
Value = archived
Field :
Name = memberarchive
Data Type = mksItem
Item :
Id = / Projects / FCA / VF_VSIM_2022 / Software Development / Eng / Bootload / CBD2000646_D00 / Misc / HexView / license.liz
Context = NULL
Model Type = si.Archive
Field :
Name = name
Data Type = wchar_t*
Value = / Projects / FCA / VF_VSIM_2022 / Software Development / Eng / Tools / HexView / license.liz
Field :
Name = memberrev
Data Type = mksItem
Item :
Id = 1.4
Context = NULL
Model Type = si.Revision
Field :
Name = cpid
Data Type = mksItem
Item :
Id = 349176 : 3
Context = NULL
Model Type = si.ChangePackage
*/