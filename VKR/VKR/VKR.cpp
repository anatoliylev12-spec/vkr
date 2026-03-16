#include "VKR.h"

using namespace System;
using namespace System::Windows::Forms;

[STAThreadAttribute]
int main(array<String^>^ args)
{
    Application::EnableVisualStyles();
    Application::SetCompatibleTextRenderingDefault(false);

    try
    {
        Application::Run(gcnew ModbusMasterForm());
    }
    catch (Exception^ ex)
    {
        MessageBox::Show("Ошибка: " + ex->Message, "Ошибка",
            MessageBoxButtons::OK, MessageBoxIcon::Error);
    }

    return 0;
}