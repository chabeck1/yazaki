#include "pch.h"

using namespace System;

int main(array<System::String ^> ^args)
{
    int State = 0;
    System::Xml::XmlTextReader^ reader = gcnew System::Xml::XmlTextReader(args[0]);
    //System::Collections::Generic::SortedSet<System::String^>^ Rules = gcnew System::Collections::Generic::SortedSet<System::String^>();
    //System::Collections::Generic::   
    System::String^ filename;
    while (reader->Read()) {
        switch (reader->NodeType) {
        case System::Xml::XmlNodeType::Element:
            //System::Console::WriteLine("<{0}>", reader->Name);
            if (reader->Name == "qav:group")
            {
                //Rules->Add(reader->Value);
            }
            if (reader->Name == "qav:filename")
            {
                //System::Console::WriteLine (reader->GetAttribute("fileid"));
                reader->Read();
                reader->Read();
                filename = reader->Value;
            }
            if (reader->Name == "qav:value")
            {
                //System::String^ fileid = reader->GetAttribute("fileid");
                System::String^ msgid = reader->GetAttribute("msgid");
                reader->Read();
                if (reader->Value != "0")
                    System::Console::WriteLine("{0}\t{1}\t{2}", msgid, filename, reader->Value);
            }
        case System::Xml::XmlNodeType::Text:
            //System::Console::WriteLine(reader->Value);
            break;
        case System::Xml::XmlNodeType::CDATA:
            break;
        case System::Xml::XmlNodeType::ProcessingInstruction:
            //System::Console::WriteLine("<?{0} {1}?>", reader->Name, reader->Value);
            break;
        case System::Xml::XmlNodeType::Comment:
            //System::Console::WriteLine("<!--{0}-->", reader->Value);
            break;
        case System::Xml::XmlNodeType::XmlDeclaration:
            //System::Console::WriteLine("<?xml version='1.0'?>");
            break;
        case System::Xml::XmlNodeType::Document:
            break;
        case System::Xml::XmlNodeType::DocumentType:
            //System::Console::WriteLine("<!DOCTYPE {0} [{1}]", reader->Name, reader->Value);
            break;
        case System::Xml::XmlNodeType::EntityReference:
            //System::Console::WriteLine(reader->Name);
            break;
        case System::Xml::XmlNodeType::EndElement:
            //System::Console::WriteLine("</{0}>", reader->Name);
            break;
        }
    }
    reader->Close();
}

#if 0
    rdr["qav:compliance matrix items"];
    while (rdr->MoveToElement)
    while (rdr->Read())
    {
        if (rdr->Name == "qav:filename")
        {
        }
        if (rdr->NodeType == System::Xml::System::Xml::XmlNodeType:::CDATA)
        {
            System::Diagnostics::Debug::WriteLineLine(rdr->Value);
        }
    }
    return 0;
}
#endif