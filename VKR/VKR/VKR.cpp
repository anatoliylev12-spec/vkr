// VKR.cpp
#include "VKR.h"

using namespace System;
using namespace System::Windows::Forms;
using namespace System::Timers;

// ============================================================================
// Реализация RegistersViewerForm
// ============================================================================

RegistersViewerForm::RegistersViewerForm(void)
{
    InitializeComponent();
    InitializeParameters();
    InitializeTimer();
    isRegistersPanelExpanded = false;
}

RegistersViewerForm::~RegistersViewerForm()
{
    if (components)
    {
        delete components;
    }
    if (updateTimer != nullptr)
    {
        updateTimer->Stop();
        delete updateTimer;
    }
}

void RegistersViewerForm::InitializeParameters()
{
    parametersList = gcnew List<ParameterInfo^>();
    parametersList->Add(gcnew ParameterInfo(L"Плотность", 257, L"кг/м³"));
    parametersList->Add(gcnew ParameterInfo(L"Температура", 259, L"°C"));
    parametersList->Add(gcnew ParameterInfo(L"Uout", 264, L"В"));
    parametersList->Add(gcnew ParameterInfo(L"Период А", 286, L"мкс"));
    parametersList->Add(gcnew ParameterInfo(L"Ku", 285, L""));
    parametersList->Add(gcnew ParameterInfo(L"Плотность NC", 288, L"кг/м³"));
    parametersList->Add(gcnew ParameterInfo(L"Добротность", 265, L""));

    changedRegisters = gcnew Dictionary<int, RegisterChange^>();
}

void RegistersViewerForm::InitializeTimer()
{
    updateTimer = gcnew System::Timers::Timer();
    updateTimer->Interval = 4000;
    updateTimer->AutoReset = true;
    updateTimer->Elapsed += gcnew System::Timers::ElapsedEventHandler(this, &RegistersViewerForm::OnTimerElapsed);
    isAutoUpdating = false;
}

void RegistersViewerForm::OnTimerElapsed(Object^ sender, System::Timers::ElapsedEventArgs^ e)
{
    if (this->InvokeRequired)
    {
        this->BeginInvoke(gcnew Action(this, &RegistersViewerForm::AutoUpdateData));
    }
    else
    {
        AutoUpdateData();
    }
}

void RegistersViewerForm::AutoUpdateData()
{
    if (OnRefreshData != nullptr && isAutoUpdating)
    {
        lastUpdateTime = DateTime::Now;
        labelAutoUpdateStatus->Text = String::Format(L"Автообновление: активно (последнее: {0:HH:mm:ss})", lastUpdateTime);
        OnRefreshData();
    }
}

void RegistersViewerForm::InitializeComponent()
{
    this->dataGridViewRegisters = (gcnew System::Windows::Forms::DataGridView());
    this->dataGridViewParameters = (gcnew System::Windows::Forms::DataGridView());
    this->buttonExport = (gcnew System::Windows::Forms::Button());
    this->buttonRefresh = (gcnew System::Windows::Forms::Button());
    this->buttonStartAutoUpdate = (gcnew System::Windows::Forms::Button());
    this->buttonStopAutoUpdate = (gcnew System::Windows::Forms::Button());
    this->buttonToggleRegisters = (gcnew System::Windows::Forms::Button());
    this->buttonSendChanges = (gcnew System::Windows::Forms::Button());
    this->labelAutoUpdateStatus = (gcnew System::Windows::Forms::Label());
    this->groupBoxRegisters = (gcnew System::Windows::Forms::GroupBox());
    this->groupBoxParameters = (gcnew System::Windows::Forms::GroupBox());
    this->saveFileDialog = (gcnew System::Windows::Forms::SaveFileDialog());
    (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridViewRegisters))->BeginInit();
    (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridViewParameters))->BeginInit();
    this->groupBoxRegisters->SuspendLayout();
    this->groupBoxParameters->SuspendLayout();
    this->SuspendLayout();

    // dataGridViewRegisters
    this->dataGridViewRegisters->AllowUserToAddRows = false;
    this->dataGridViewRegisters->AllowUserToDeleteRows = false;
    this->dataGridViewRegisters->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
    this->dataGridViewRegisters->Dock = System::Windows::Forms::DockStyle::Fill;
    this->dataGridViewRegisters->Location = System::Drawing::Point(3, 22);
    this->dataGridViewRegisters->Name = L"dataGridViewRegisters";
    this->dataGridViewRegisters->ReadOnly = false;
    this->dataGridViewRegisters->RowHeadersWidth = 62;
    this->dataGridViewRegisters->Size = System::Drawing::Size(600, 200);
    this->dataGridViewRegisters->TabIndex = 0;
    this->dataGridViewRegisters->CellValueChanged += gcnew System::Windows::Forms::DataGridViewCellEventHandler(this, &RegistersViewerForm::dataGridViewRegisters_CellValueChanged);

    // dataGridViewParameters
    this->dataGridViewParameters->AllowUserToAddRows = false;
    this->dataGridViewParameters->AllowUserToDeleteRows = false;
    this->dataGridViewParameters->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
    this->dataGridViewParameters->Dock = System::Windows::Forms::DockStyle::Fill;
    this->dataGridViewParameters->Location = System::Drawing::Point(3, 22);
    this->dataGridViewParameters->Name = L"dataGridViewParameters";
    this->dataGridViewParameters->ReadOnly = true;
    this->dataGridViewParameters->RowHeadersWidth = 62;
    this->dataGridViewParameters->Size = System::Drawing::Size(350, 450);
    this->dataGridViewParameters->TabIndex = 1;

    // buttonExport
    this->buttonExport->Location = System::Drawing::Point(12, 518);
    this->buttonExport->Name = L"buttonExport";
    this->buttonExport->Size = System::Drawing::Size(100, 35);
    this->buttonExport->TabIndex = 1;
    this->buttonExport->Text = L"Экспорт в CSV";
    this->buttonExport->UseVisualStyleBackColor = true;
    this->buttonExport->Click += gcnew System::EventHandler(this, &RegistersViewerForm::buttonExport_Click);

    // buttonRefresh
    this->buttonRefresh->Location = System::Drawing::Point(118, 518);
    this->buttonRefresh->Name = L"buttonRefresh";
    this->buttonRefresh->Size = System::Drawing::Size(100, 35);
    this->buttonRefresh->TabIndex = 2;
    this->buttonRefresh->Text = L"Обновить";
    this->buttonRefresh->UseVisualStyleBackColor = true;
    this->buttonRefresh->Click += gcnew System::EventHandler(this, &RegistersViewerForm::buttonRefresh_Click);

    // buttonStartAutoUpdate
    this->buttonStartAutoUpdate->Location = System::Drawing::Point(224, 518);
    this->buttonStartAutoUpdate->Name = L"buttonStartAutoUpdate";
    this->buttonStartAutoUpdate->Size = System::Drawing::Size(120, 35);
    this->buttonStartAutoUpdate->TabIndex = 3;
    this->buttonStartAutoUpdate->Text = L"Автообновление ВКЛ";
    this->buttonStartAutoUpdate->UseVisualStyleBackColor = true;
    this->buttonStartAutoUpdate->Click += gcnew System::EventHandler(this, &RegistersViewerForm::buttonStartAutoUpdate_Click);

    // buttonStopAutoUpdate
    this->buttonStopAutoUpdate->Enabled = false;
    this->buttonStopAutoUpdate->Location = System::Drawing::Point(350, 518);
    this->buttonStopAutoUpdate->Name = L"buttonStopAutoUpdate";
    this->buttonStopAutoUpdate->Size = System::Drawing::Size(120, 35);
    this->buttonStopAutoUpdate->TabIndex = 4;
    this->buttonStopAutoUpdate->Text = L"Автообновление ВЫКЛ";
    this->buttonStopAutoUpdate->UseVisualStyleBackColor = true;
    this->buttonStopAutoUpdate->Click += gcnew System::EventHandler(this, &RegistersViewerForm::buttonStopAutoUpdate_Click);

    // buttonToggleRegisters
    this->buttonToggleRegisters->Location = System::Drawing::Point(476, 518);
    this->buttonToggleRegisters->Name = L"buttonToggleRegisters";
    this->buttonToggleRegisters->Size = System::Drawing::Size(140, 35);
    this->buttonToggleRegisters->TabIndex = 5;
    this->buttonToggleRegisters->Text = L"Развернуть таблицу регистров";
    this->buttonToggleRegisters->UseVisualStyleBackColor = true;
    this->buttonToggleRegisters->Click += gcnew System::EventHandler(this, &RegistersViewerForm::buttonToggleRegisters_Click);

    // buttonSendChanges
    this->buttonSendChanges->Enabled = false;
    this->buttonSendChanges->Location = System::Drawing::Point(622, 518);
    this->buttonSendChanges->Name = L"buttonSendChanges";
    this->buttonSendChanges->Size = System::Drawing::Size(120, 35);
    this->buttonSendChanges->TabIndex = 6;
    this->buttonSendChanges->Text = L"Отправить изменения";
    this->buttonSendChanges->UseVisualStyleBackColor = true;
    this->buttonSendChanges->Click += gcnew System::EventHandler(this, &RegistersViewerForm::buttonSendChanges_Click);

    // labelAutoUpdateStatus
    this->labelAutoUpdateStatus->AutoSize = true;
    this->labelAutoUpdateStatus->Location = System::Drawing::Point(12, 560);
    this->labelAutoUpdateStatus->Name = L"labelAutoUpdateStatus";
    this->labelAutoUpdateStatus->Size = System::Drawing::Size(170, 20);
    this->labelAutoUpdateStatus->TabIndex = 5;
    this->labelAutoUpdateStatus->Text = L"Автообновление: выключено";

    // groupBoxRegisters
    this->groupBoxRegisters->Controls->Add(this->dataGridViewRegisters);
    this->groupBoxRegisters->Location = System::Drawing::Point(12, 12);
    this->groupBoxRegisters->Name = L"groupBoxRegisters";
    this->groupBoxRegisters->Size = System::Drawing::Size(606, 250);
    this->groupBoxRegisters->TabIndex = 6;
    this->groupBoxRegisters->TabStop = false;
    this->groupBoxRegisters->Text = L"Все регистры (32-битные значения) [СВЕРНУТО]";

    // groupBoxParameters
    this->groupBoxParameters->Controls->Add(this->dataGridViewParameters);
    this->groupBoxParameters->Location = System::Drawing::Point(624, 12);
    this->groupBoxParameters->Name = L"groupBoxParameters";
    this->groupBoxParameters->Size = System::Drawing::Size(356, 475);
    this->groupBoxParameters->TabIndex = 7;
    this->groupBoxParameters->TabStop = false;
    this->groupBoxParameters->Text = L"Контролируемые параметры";

    // RegistersViewerForm
    this->AutoScaleDimensions = System::Drawing::SizeF(9, 20);
    this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
    this->ClientSize = System::Drawing::Size(992, 590);
    this->Controls->Add(this->groupBoxParameters);
    this->Controls->Add(this->groupBoxRegisters);
    this->Controls->Add(this->labelAutoUpdateStatus);
    this->Controls->Add(this->buttonSendChanges);
    this->Controls->Add(this->buttonToggleRegisters);
    this->Controls->Add(this->buttonStopAutoUpdate);
    this->Controls->Add(this->buttonStartAutoUpdate);
    this->Controls->Add(this->buttonRefresh);
    this->Controls->Add(this->buttonExport);
    this->Name = L"RegistersViewerForm";
    this->Text = L"Регистры устройства (32-битные значения, Little-Endian)";
    this->FormClosing += gcnew System::Windows::Forms::FormClosingEventHandler(this, &RegistersViewerForm::RegistersViewerForm_FormClosing);
    this->Load += gcnew System::EventHandler(this, &RegistersViewerForm::RegistersViewerForm_Load);

    (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridViewRegisters))->EndInit();
    (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridViewParameters))->EndInit();
    this->groupBoxRegisters->ResumeLayout(false);
    this->groupBoxParameters->ResumeLayout(false);
    this->ResumeLayout(false);
    this->PerformLayout();

    originalFormHeight = this->Height;
    originalRegistersGroupHeight = groupBoxRegisters->Height;
    CollapseRegistersPanel();
}

void RegistersViewerForm::UpdateRegistersData(List<System::Tuple<int, float, int32_t>^>^ registersData)
{
    allRegistersData = registersData;
    UpdateFullRegistersTable(registersData);
    UpdateParametersTable(registersData);

    if (!isAutoUpdating)
    {
        changedRegisters->Clear();
        buttonSendChanges->Enabled = false;
    }
}

void RegistersViewerForm::UpdateFullRegistersTable(List<System::Tuple<int, float, int32_t>^>^ registersData)
{
    int firstDisplayedScrollRow = 0;
    if (dataGridViewRegisters->FirstDisplayedScrollingRowIndex >= 0)
        firstDisplayedScrollRow = dataGridViewRegisters->FirstDisplayedScrollingRowIndex;

    dataGridViewRegisters->Rows->Clear();
    dataGridViewRegisters->Columns->Clear();
    dataGridViewRegisters->Columns->Add(L"Index", L"№ п/п");
    dataGridViewRegisters->Columns->Add(L"RegisterAddr", L"Адрес (dec)");
    dataGridViewRegisters->Columns->Add(L"ValueHex", L"Значение (hex)");
    dataGridViewRegisters->Columns->Add(L"ValueInt32", L"Значение (int32)");
    dataGridViewRegisters->Columns->Add(L"ValueFloat", L"Значение (float)");

    dataGridViewRegisters->Columns[0]->Width = 60;
    dataGridViewRegisters->Columns[1]->Width = 120;
    dataGridViewRegisters->Columns[2]->Width = 120;
    dataGridViewRegisters->Columns[3]->Width = 120;
    dataGridViewRegisters->Columns[4]->Width = 120;

    for each (auto item in registersData)
    {
        int index = item->Item1;
        float floatValue = item->Item2;
        int32_t intValue = item->Item3;

        int startAddr = index * 2;
        int endAddr = index * 2 + 1;
        String^ addrStr = String::Format(L"{0}-{1} (0x{2:X4}-0x{3:X4})", startAddr, endAddr, startAddr, endAddr);

        uint32_t unsignedValue = static_cast<uint32_t>(intValue);
        String^ hexStr = String::Format(L"0x{0:X8}", unsignedValue);

        DataGridViewRow^ row = gcnew DataGridViewRow();
        row->CreateCells(dataGridViewRegisters);
        row->Cells[0]->Value = (index + 1).ToString();
        row->Cells[1]->Value = addrStr;
        row->Cells[2]->Value = hexStr;
        row->Cells[3]->Value = intValue.ToString();
        row->Cells[4]->Value = floatValue.ToString(L"F6");

        if (changedRegisters->ContainsKey(index))
        {
            row->DefaultCellStyle->BackColor = Color::LightYellow;
        }

        dataGridViewRegisters->Rows->Add(row);
    }

    if (firstDisplayedScrollRow > 0 && firstDisplayedScrollRow < dataGridViewRegisters->Rows->Count)
    {
        dataGridViewRegisters->FirstDisplayedScrollingRowIndex = firstDisplayedScrollRow;
    }

    dataGridViewRegisters->AutoResizeColumns(DataGridViewAutoSizeColumnsMode::AllCells);
    dataGridViewRegisters->AutoSizeRowsMode = DataGridViewAutoSizeRowsMode::AllCells;
}

void RegistersViewerForm::UpdateParametersTable(List<System::Tuple<int, float, int32_t>^>^ registersData)
{
    dataGridViewParameters->Rows->Clear();
    dataGridViewParameters->Columns->Clear();
    dataGridViewParameters->Columns->Add(L"DateTime", L"Дата/Время");
    dataGridViewParameters->Columns->Add(L"Parameter", L"Параметр");
    dataGridViewParameters->Columns->Add(L"RegisterAddr", L"Адрес");
    dataGridViewParameters->Columns->Add(L"Value", L"Значение");
    dataGridViewParameters->Columns->Add(L"Unit", L"Ед. изм.");

    dataGridViewParameters->Columns[0]->Width = 120;
    dataGridViewParameters->Columns[1]->Width = 100;
    dataGridViewParameters->Columns[2]->Width = 80;
    dataGridViewParameters->Columns[3]->Width = 100;
    dataGridViewParameters->Columns[4]->Width = 80;

    Dictionary<int, float>^ valueDict = gcnew Dictionary<int, float>();
    for each (auto item in registersData)
    {
        valueDict[item->Item1] = item->Item2;
    }

    String^ currentTime = DateTime::Now.ToString(L"dd.MM.yyyy HH:mm:ss");

    for each (ParameterInfo ^ param in parametersList)
    {
        float value = 0.0f;
        if (valueDict->ContainsKey(param->RegisterAddress))
        {
            value = valueDict[param->RegisterAddress];
        }

        int startAddr = param->RegisterAddress * 2;
        String^ addrStr = String::Format(L"0x{0:X4} ({0})", startAddr);

        dataGridViewParameters->Rows->Add(currentTime, param->Name, addrStr, value.ToString(L"F6"), param->Unit);
    }

    dataGridViewParameters->AutoResizeColumns(DataGridViewAutoSizeColumnsMode::AllCells);
    dataGridViewParameters->AutoSizeRowsMode = DataGridViewAutoSizeRowsMode::AllCells;
}

void RegistersViewerForm::CollapseRegistersPanel()
{
    groupBoxRegisters->Height = 50;
    dataGridViewRegisters->Visible = false;
    groupBoxRegisters->Text = L"Все регистры (32-битные значения) [СВЕРНУТО]";
    buttonToggleRegisters->Text = L"Развернуть таблицу регистров";
    isRegistersPanelExpanded = false;
    groupBoxParameters->Location = System::Drawing::Point(624, 12);
    groupBoxParameters->Height = this->Height - 120;
}

void RegistersViewerForm::ExpandRegistersPanel()
{
    groupBoxRegisters->Height = 250;
    dataGridViewRegisters->Visible = true;
    groupBoxRegisters->Text = L"Все регистры (32-битные значения) [РАЗВЕРНУТО]";
    buttonToggleRegisters->Text = L"Свернуть таблицу регистров";
    isRegistersPanelExpanded = true;
    groupBoxParameters->Location = System::Drawing::Point(624, 12);
    groupBoxParameters->Height = 475;
}

void RegistersViewerForm::dataGridViewRegisters_CellValueChanged(Object^ sender, DataGridViewCellEventArgs^ e)
{
    if (e->RowIndex >= 0 && e->ColumnIndex >= 3 && e->ColumnIndex <= 4)
    {
        int registerIndex = e->RowIndex;
        String^ newValueStr = dataGridViewRegisters->Rows[e->RowIndex]->Cells[e->ColumnIndex]->Value->ToString();

        try
        {
            int32_t newIntValue;
            float newFloatValue;
            uint32_t newUintValue;

            if (e->ColumnIndex == 3)
            {
                newIntValue = Convert::ToInt32(newValueStr);
                newFloatValue = static_cast<float>(newIntValue);
                newUintValue = static_cast<uint32_t>(newIntValue);
                dataGridViewRegisters->Rows[e->RowIndex]->Cells[4]->Value = newFloatValue.ToString(L"F6");
            }
            else
            {
                newFloatValue = Convert::ToSingle(newValueStr);
                newIntValue = static_cast<int32_t>(newFloatValue);
                newUintValue = static_cast<uint32_t>(newIntValue);
                dataGridViewRegisters->Rows[e->RowIndex]->Cells[3]->Value = newIntValue.ToString();
            }

            dataGridViewRegisters->Rows[e->RowIndex]->Cells[2]->Value = String::Format(L"0x{0:X8}", newUintValue);

            RegisterChange^ change = gcnew RegisterChange(registerIndex, newUintValue, newFloatValue);
            changedRegisters[registerIndex] = change;
            dataGridViewRegisters->Rows[e->RowIndex]->DefaultCellStyle->BackColor = Color::LightYellow;
            buttonSendChanges->Enabled = true;
        }
        catch (Exception^ ex)
        {
            MessageBox::Show(L"Ошибка ввода значения: " + ex->Message, L"Ошибка", MessageBoxButtons::OK, MessageBoxIcon::Warning);
            UpdateFullRegistersTable(allRegistersData);
        }
    }
}

void RegistersViewerForm::buttonExport_Click(Object^ sender, EventArgs^ e)
{
    if (dataGridViewRegisters->Rows->Count == 0)
    {
        MessageBox::Show(L"Нет данных для экспорта!", L"Предупреждение", MessageBoxButtons::OK, MessageBoxIcon::Warning);
        return;
    }

    saveFileDialog->Filter = L"CSV files (*.csv)|*.csv|All files (*.*)|*.*";
    saveFileDialog->FilterIndex = 1;
    saveFileDialog->RestoreDirectory = true;
    saveFileDialog->FileName = L"registers_export.csv";

    if (saveFileDialog->ShowDialog() == System::Windows::Forms::DialogResult::OK)
    {
        try
        {
            System::IO::StreamWriter^ sw = gcnew System::IO::StreamWriter(saveFileDialog->FileName);

            sw->WriteLine(L"=== ВСЕ РЕГИСТРЫ ===");
            for (int i = 0; i < dataGridViewRegisters->Columns->Count; i++)
            {
                sw->Write(dataGridViewRegisters->Columns[i]->HeaderText);
                if (i < dataGridViewRegisters->Columns->Count - 1) sw->Write(L";");
            }
            sw->WriteLine();

            for (int i = 0; i < dataGridViewRegisters->Rows->Count; i++)
            {
                for (int j = 0; j < dataGridViewRegisters->Columns->Count; j++)
                {
                    String^ cellValue = dataGridViewRegisters->Rows[i]->Cells[j]->Value != nullptr ?
                        dataGridViewRegisters->Rows[i]->Cells[j]->Value->ToString() : L"";
                    cellValue = cellValue->Replace(L"\"", L"\"\"");
                    if (cellValue->Contains(L";") || cellValue->Contains(L"\""))
                        sw->Write(L"\"" + cellValue + L"\"");
                    else
                        sw->Write(cellValue);
                    if (j < dataGridViewRegisters->Columns->Count - 1) sw->Write(L";");
                }
                sw->WriteLine();
            }

            sw->WriteLine();
            sw->WriteLine(L"=== КОНТРОЛИРУЕМЫЕ ПАРАМЕТРЫ ===");
            for (int i = 0; i < dataGridViewParameters->Columns->Count; i++)
            {
                sw->Write(dataGridViewParameters->Columns[i]->HeaderText);
                if (i < dataGridViewParameters->Columns->Count - 1) sw->Write(L";");
            }
            sw->WriteLine();

            for (int i = 0; i < dataGridViewParameters->Rows->Count; i++)
            {
                for (int j = 0; j < dataGridViewParameters->Columns->Count; j++)
                {
                    String^ cellValue = dataGridViewParameters->Rows[i]->Cells[j]->Value != nullptr ?
                        dataGridViewParameters->Rows[i]->Cells[j]->Value->ToString() : L"";
                    cellValue = cellValue->Replace(L"\"", L"\"\"");
                    if (cellValue->Contains(L";") || cellValue->Contains(L"\""))
                        sw->Write(L"\"" + cellValue + L"\"");
                    else
                        sw->Write(cellValue);
                    if (j < dataGridViewParameters->Columns->Count - 1) sw->Write(L";");
                }
                sw->WriteLine();
            }

            sw->Close();
            MessageBox::Show(L"Данные успешно экспортированы в " + saveFileDialog->FileName, L"Экспорт завершен",
                MessageBoxButtons::OK, MessageBoxIcon::Information);
        }
        catch (Exception^ ex)
        {
            MessageBox::Show(L"Ошибка при экспорте: " + ex->Message, L"Ошибка", MessageBoxButtons::OK, MessageBoxIcon::Error);
        }
    }
}

void RegistersViewerForm::buttonRefresh_Click(Object^ sender, EventArgs^ e)
{
    if (OnRefreshData != nullptr)
    {
        OnRefreshData();
        lastUpdateTime = DateTime::Now;
        if (isAutoUpdating)
        {
            labelAutoUpdateStatus->Text = String::Format(L"Автообновление: активно (последнее: {0:HH:mm:ss})", lastUpdateTime);
        }
    }
}

void RegistersViewerForm::buttonStartAutoUpdate_Click(Object^ sender, EventArgs^ e)
{
    if (OnRefreshData != nullptr)
    {
        isAutoUpdating = true;
        updateTimer->Start();
        buttonStartAutoUpdate->Enabled = false;
        buttonStopAutoUpdate->Enabled = true;
        labelAutoUpdateStatus->Text = L"Автообновление: активно (обновление каждые 4 сек)";
        lastUpdateTime = DateTime::Now;
        OnRefreshData();
    }
    else
    {
        MessageBox::Show(L"Не задана функция обновления данных!", L"Ошибка", MessageBoxButtons::OK, MessageBoxIcon::Warning);
    }
}

void RegistersViewerForm::buttonStopAutoUpdate_Click(Object^ sender, EventArgs^ e)
{
    isAutoUpdating = false;
    updateTimer->Stop();
    buttonStartAutoUpdate->Enabled = true;
    buttonStopAutoUpdate->Enabled = false;
    labelAutoUpdateStatus->Text = L"Автообновление: выключено";
}

void RegistersViewerForm::buttonToggleRegisters_Click(Object^ sender, EventArgs^ e)
{
    if (isRegistersPanelExpanded)
        CollapseRegistersPanel();
    else
        ExpandRegistersPanel();
}

void RegistersViewerForm::buttonSendChanges_Click(Object^ sender, EventArgs^ e)
{
    if (changedRegisters->Count == 0)
    {
        MessageBox::Show(L"Нет изменений для отправки!", L"Информация", MessageBoxButtons::OK, MessageBoxIcon::Information);
        return;
    }

    if (OnWriteRegisters != nullptr)
    {
        List<RegisterChange^>^ changesList = gcnew List<RegisterChange^>();
        for each (auto kvp in changedRegisters)
        {
            changesList->Add(kvp.Value);
        }

        OnWriteRegisters(changesList);
        changedRegisters->Clear();
        buttonSendChanges->Enabled = false;

        if (OnRefreshData != nullptr)
        {
            OnRefreshData();
        }
    }
}

void RegistersViewerForm::RegistersViewerForm_Load(Object^ sender, EventArgs^ e)
{
    CollapseRegistersPanel();
}

void RegistersViewerForm::RegistersViewerForm_FormClosing(Object^ sender, FormClosingEventArgs^ e)
{
    if (updateTimer != nullptr)
    {
        updateTimer->Stop();
    }
}

// ============================================================================
// Реализация ModbusMasterForm
// ============================================================================

ModbusMasterForm::ModbusMasterForm(void)
{
    InitializeComponent();
    InitializeCustomComponent();
    hSerial = INVALID_HANDLE_VALUE;
    baudRate = CBR_9600;
    byteSize = 8;
    stopBits = ONESTOPBIT;
    parity = NOPARITY;
    allRegistersData = gcnew List<System::Tuple<int, float, int32_t>^>();
    isViewerFormOpen = false;
}

ModbusMasterForm::~ModbusMasterForm()
{
    if (components)
    {
        delete components;
    }
    if (hSerial != INVALID_HANDLE_VALUE)
    {
        ClosePort();
    }
}

void ModbusMasterForm::InitializeCustomComponent()
{
    numericSlaveID->Minimum = MIN_SLAVE_ID;
    numericSlaveID->Maximum = MAX_SLAVE_ID;
    numericSlaveID->Value = 1;

    buttonOpenViewer = gcnew System::Windows::Forms::Button();
    buttonOpenViewer->Location = System::Drawing::Point(548, 30);
    buttonOpenViewer->Name = L"buttonOpenViewer";
    buttonOpenViewer->Size = System::Drawing::Size(150, 30);
    buttonOpenViewer->TabIndex = 4;
    buttonOpenViewer->Text = L"Открыть монитор регистров";
    buttonOpenViewer->UseVisualStyleBackColor = true;
    buttonOpenViewer->Enabled = false;
    buttonOpenViewer->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonOpenViewer_Click);
    this->groupBoxPort->Controls->Add(buttonOpenViewer);

    isViewerFormOpen = false;
}

void ModbusMasterForm::InitializeComponent()
{
    this->comboBoxPorts = (gcnew System::Windows::Forms::ComboBox());
    this->buttonRefresh = (gcnew System::Windows::Forms::Button());
    this->buttonConnect = (gcnew System::Windows::Forms::Button());
    this->buttonDisconnect = (gcnew System::Windows::Forms::Button());
    this->buttonScanDevices = (gcnew System::Windows::Forms::Button());
    this->textBoxLog = (gcnew System::Windows::Forms::TextBox());
    this->labelStatus = (gcnew System::Windows::Forms::Label());
    this->groupBoxPort = (gcnew System::Windows::Forms::GroupBox());
    this->labelSlaveID = (gcnew System::Windows::Forms::Label());
    this->numericSlaveID = (gcnew System::Windows::Forms::NumericUpDown());
    this->checkBoxShowAll = (gcnew System::Windows::Forms::CheckBox());
    this->buttonClearLog = (gcnew System::Windows::Forms::Button());
    this->labelFoundDevices = (gcnew System::Windows::Forms::Label());
    this->textBoxScanResult = (gcnew System::Windows::Forms::TextBox());
    this->groupBoxPort->SuspendLayout();
    (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->numericSlaveID))->BeginInit();
    this->SuspendLayout();
    // 
    // comboBoxPorts
    // 
    this->comboBoxPorts->DropDownStyle = System::Windows::Forms::ComboBoxStyle::DropDownList;
    this->comboBoxPorts->FormattingEnabled = true;
    this->comboBoxPorts->Location = System::Drawing::Point(15, 30);
    this->comboBoxPorts->Name = L"comboBoxPorts";
    this->comboBoxPorts->Size = System::Drawing::Size(150, 28);
    this->comboBoxPorts->TabIndex = 0;
    // 
    // buttonRefresh
    // 
    this->buttonRefresh->Location = System::Drawing::Point(175, 30);
    this->buttonRefresh->Name = L"buttonRefresh";
    this->buttonRefresh->Size = System::Drawing::Size(100, 30);
    this->buttonRefresh->TabIndex = 1;
    this->buttonRefresh->Text = L"Обновить порты";
    this->buttonRefresh->UseVisualStyleBackColor = true;
    this->buttonRefresh->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonRefresh_Click);
    // 
    // buttonConnect
    // 
    this->buttonConnect->Location = System::Drawing::Point(285, 30);
    this->buttonConnect->Name = L"buttonConnect";
    this->buttonConnect->Size = System::Drawing::Size(100, 30);
    this->buttonConnect->TabIndex = 2;
    this->buttonConnect->Text = L"Подключить";
    this->buttonConnect->UseVisualStyleBackColor = true;
    this->buttonConnect->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonConnect_Click);
    // 
    // buttonDisconnect
    // 
    this->buttonDisconnect->Enabled = false;
    this->buttonDisconnect->Location = System::Drawing::Point(395, 30);
    this->buttonDisconnect->Name = L"buttonDisconnect";
    this->buttonDisconnect->Size = System::Drawing::Size(100, 30);
    this->buttonDisconnect->TabIndex = 3;
    this->buttonDisconnect->Text = L"Отключить";
    this->buttonDisconnect->UseVisualStyleBackColor = true;
    this->buttonDisconnect->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonDisconnect_Click);
    // 
    // buttonScanDevices
    // 
    this->buttonScanDevices->Location = System::Drawing::Point(15, 95);
    this->buttonScanDevices->Name = L"buttonScanDevices";
    this->buttonScanDevices->Size = System::Drawing::Size(150, 30);
    this->buttonScanDevices->TabIndex = 4;
    this->buttonScanDevices->Text = L"Поиск устройств";
    this->buttonScanDevices->UseVisualStyleBackColor = true;
    this->buttonScanDevices->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonScanDevices_Click);
    // 
    // textBoxLog
    // 
    this->textBoxLog->Location = System::Drawing::Point(12, 200);
    this->textBoxLog->Multiline = true;
    this->textBoxLog->Name = L"textBoxLog";
    this->textBoxLog->ReadOnly = true;
    this->textBoxLog->ScrollBars = System::Windows::Forms::ScrollBars::Vertical;
    this->textBoxLog->Size = System::Drawing::Size(760, 300);
    this->textBoxLog->TabIndex = 5;
    // 
    // labelStatus
    // 
    this->labelStatus->AutoSize = true;
    this->labelStatus->Location = System::Drawing::Point(12, 515);
    this->labelStatus->Name = L"labelStatus";
    this->labelStatus->Size = System::Drawing::Size(148, 20);
    this->labelStatus->TabIndex = 6;
    this->labelStatus->Text = L"Статус: Отключен";
    // 
    // groupBoxPort
    // 
    this->groupBoxPort->Controls->Add(this->labelSlaveID);
    this->groupBoxPort->Controls->Add(this->numericSlaveID);
    this->groupBoxPort->Controls->Add(this->checkBoxShowAll);
    this->groupBoxPort->Controls->Add(this->comboBoxPorts);
    this->groupBoxPort->Controls->Add(this->buttonScanDevices);
    this->groupBoxPort->Controls->Add(this->buttonRefresh);
    this->groupBoxPort->Controls->Add(this->buttonDisconnect);
    this->groupBoxPort->Controls->Add(this->buttonConnect);
    this->groupBoxPort->Location = System::Drawing::Point(12, 12);
    this->groupBoxPort->Name = L"groupBoxPort";
    this->groupBoxPort->Size = System::Drawing::Size(760, 170);
    this->groupBoxPort->TabIndex = 7;
    this->groupBoxPort->TabStop = false;
    this->groupBoxPort->Text = L"Управление портом (RS-485, 9600 бод)";
    // 
    // labelSlaveID
    // 
    this->labelSlaveID->AutoSize = true;
    this->labelSlaveID->Location = System::Drawing::Point(505, 70);
    this->labelSlaveID->Name = L"labelSlaveID";
    this->labelSlaveID->Size = System::Drawing::Size(120, 20);
    this->labelSlaveID->TabIndex = 12;
    this->labelSlaveID->Text = L"ID устройства:";
    // 
    // numericSlaveID
    // 
    this->numericSlaveID->Location = System::Drawing::Point(505, 95);
    this->numericSlaveID->Maximum = System::Decimal(gcnew cli::array< System::Int32 >(4) { 247, 0, 0, 0 });
    this->numericSlaveID->Minimum = System::Decimal(gcnew cli::array< System::Int32 >(4) { 1, 0, 0, 0 });
    this->numericSlaveID->Name = L"numericSlaveID";
    this->numericSlaveID->Size = System::Drawing::Size(60, 26);
    this->numericSlaveID->TabIndex = 11;
    this->numericSlaveID->Value = System::Decimal(gcnew cli::array< System::Int32 >(4) { 1, 0, 0, 0 });
    // 
    // checkBoxShowAll
    // 
    this->checkBoxShowAll->AutoSize = true;
    this->checkBoxShowAll->Location = System::Drawing::Point(15, 65);
    this->checkBoxShowAll->Name = L"checkBoxShowAll";
    this->checkBoxShowAll->Size = System::Drawing::Size(240, 24);
    this->checkBoxShowAll->TabIndex = 5;
    this->checkBoxShowAll->Text = L"Показывать занятые порты";
    this->checkBoxShowAll->UseVisualStyleBackColor = true;
    // 
    // buttonClearLog
    // 
    this->buttonClearLog->Location = System::Drawing::Point(700, 510);
    this->buttonClearLog->Name = L"buttonClearLog";
    this->buttonClearLog->Size = System::Drawing::Size(72, 30);
    this->buttonClearLog->TabIndex = 10;
    this->buttonClearLog->Text = L"Очистить лог";
    this->buttonClearLog->UseVisualStyleBackColor = true;
    this->buttonClearLog->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonClearLog_Click);
    // 
    // labelFoundDevices
    // 
    this->labelFoundDevices->AutoSize = true;
    this->labelFoundDevices->Location = System::Drawing::Point(12, 190);
    this->labelFoundDevices->Name = L"labelFoundDevices";
    this->labelFoundDevices->Size = System::Drawing::Size(191, 20);
    this->labelFoundDevices->TabIndex = 8;
    this->labelFoundDevices->Text = L"Найденные устройства:";
    this->labelFoundDevices->Visible = false;
    // 
    // textBoxScanResult
    // 
    this->textBoxScanResult->Location = System::Drawing::Point(12, 210);
    this->textBoxScanResult->Multiline = true;
    this->textBoxScanResult->Name = L"textBoxScanResult";
    this->textBoxScanResult->Size = System::Drawing::Size(760, 80);
    this->textBoxScanResult->TabIndex = 9;
    this->textBoxScanResult->Visible = false;
    // 
    // ModbusMasterForm
    // 
    this->AutoScaleDimensions = System::Drawing::SizeF(9, 20);
    this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
    this->ClientSize = System::Drawing::Size(784, 551);
    this->Controls->Add(this->buttonClearLog);
    this->Controls->Add(this->textBoxScanResult);
    this->Controls->Add(this->labelFoundDevices);
    this->Controls->Add(this->groupBoxPort);
    this->Controls->Add(this->labelStatus);
    this->Controls->Add(this->textBoxLog);
    this->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
        static_cast<System::Byte>(204)));
    this->Margin = System::Windows::Forms::Padding(4, 5, 4, 5);
    this->Name = L"ModbusMasterForm";
    this->Text = L"Modbus Master для RS-485 устройства (Автоматизированный режим)";
    this->Load += gcnew System::EventHandler(this, &ModbusMasterForm::Form1_Load);
    this->groupBoxPort->ResumeLayout(false);
    this->groupBoxPort->PerformLayout();
    (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->numericSlaveID))->EndInit();
    this->ResumeLayout(false);
    this->PerformLayout();

}

void ModbusMasterForm::Form1_Load(Object^ sender, EventArgs^ e)
{
    RefreshPorts();
    LogMessage(L"Программа запущена. Режим RS-485, 9600 бод.");
}

void ModbusMasterForm::buttonRefresh_Click(Object^ sender, EventArgs^ e)
{
    RefreshPorts();
}

void ModbusMasterForm::buttonConnect_Click(Object^ sender, EventArgs^ e)
{
    ConnectToPort();
}

void ModbusMasterForm::buttonDisconnect_Click(Object^ sender, EventArgs^ e)
{
    DisconnectFromPort();
}

void ModbusMasterForm::buttonScanDevices_Click(Object^ sender, EventArgs^ e)
{
    ScanForDevices();
}

void ModbusMasterForm::buttonClearLog_Click(Object^ sender, EventArgs^ e)
{
    textBoxLog->Clear();
}

void ModbusMasterForm::buttonOpenViewer_Click(Object^ sender, EventArgs^ e)
{
    OpenRegistersViewer();
}

void ModbusMasterForm::OpenRegistersViewer()
{
    if (hSerial == INVALID_HANDLE_VALUE)
    {
        MessageBox::Show(L"Порт не подключен!", L"Ошибка", MessageBoxButtons::OK, MessageBoxIcon::Warning);
        return;
    }

    LogMessage(L"Открытие монитора регистров...");

    RegistersViewerForm^ viewerForm = gcnew RegistersViewerForm();
    viewerForm->OnRefreshData = gcnew RegistersViewerForm::RefreshDataDelegate(this, &ModbusMasterForm::RefreshRegistersData);
    viewerForm->OnWriteRegisters = gcnew RegistersViewerForm::WriteRegistersDelegate(this, &ModbusMasterForm::WriteRegisters);

    ScanAllRegisters();
    viewerForm->UpdateRegistersData(allRegistersData);

    isViewerFormOpen = true;
    viewerForm->FormClosed += gcnew FormClosedEventHandler(this, &ModbusMasterForm::ViewerForm_FormClosed);
    viewerForm->Show();

    LogMessage(L"Монитор регистров открыт. Автообновление: 4 сек.");
}

void ModbusMasterForm::ViewerForm_FormClosed(Object^ sender, FormClosedEventArgs^ e)
{
    isViewerFormOpen = false;
    LogMessage(L"Монитор регистров закрыт.");
}

void ModbusMasterForm::RefreshRegistersData()
{
    if (hSerial == INVALID_HANDLE_VALUE) return;

    allRegistersData->Clear();
    ScanAllRegisters();

    for each (Form ^ form in Application::OpenForms)
    {
        RegistersViewerForm^ viewerForm = dynamic_cast<RegistersViewerForm^>(form);
        if (viewerForm != nullptr)
        {
            viewerForm->UpdateRegistersData(allRegistersData);
            break;
        }
    }
}

void ModbusMasterForm::RefreshPorts()
{
    comboBoxPorts->Items->Clear();

    if (availablePorts == nullptr)
        availablePorts = gcnew List<String^>();
    else
        availablePorts->Clear();

    bool showAll = checkBoxShowAll->Checked;

    LogMessage(L"Поиск COM-портов..." + (showAll ? L" (включая занятые)" : L" (только доступные)"));

    std::set<std::string> portSet;

    const char* registryPaths[] = {
        "HARDWARE\\DEVICEMAP\\SERIALCOMM",
        "SYSTEM\\CurrentControlSet\\Services\\Serial\\Parameters\\SerialPorts"
    };

    for (const char* registryPath : registryPaths)
    {
        HKEY hKey;
        LONG result = RegOpenKeyExA(HKEY_LOCAL_MACHINE, registryPath, 0, KEY_READ, &hKey);

        if (result == ERROR_SUCCESS)
        {
            DWORD index = 0;
            char valueName[256];
            DWORD valueNameSize;
            BYTE data[256];
            DWORD dataSize;
            DWORD type;

            while (true)
            {
                valueNameSize = sizeof(valueName);
                dataSize = sizeof(data);

                result = RegEnumValueA(hKey, index, valueName, &valueNameSize,
                    NULL, &type, data, &dataSize);

                if (result == ERROR_NO_MORE_ITEMS)
                    break;

                if (result == ERROR_SUCCESS && type == REG_SZ)
                {
                    std::string portName = "\\\\.\\" + std::string((char*)data);

                    if (showAll)
                    {
                        portSet.insert(portName);
                        LogMessage(L"Найден порт: " + StringToWString(portName));
                    }
                    else
                    {
                        HANDLE hTest = CreateFileA(portName.c_str(), GENERIC_READ | GENERIC_WRITE, 0,
                            NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);

                        if (hTest != INVALID_HANDLE_VALUE)
                        {
                            CloseHandle(hTest);
                            portSet.insert(portName);
                            LogMessage(L"Найден доступный порт: " + StringToWString(portName));
                        }
                    }
                }
                index++;
            }
            RegCloseKey(hKey);
        }
    }

    for (const auto& port : portSet)
    {
        std::string displayName = port.substr(4);
        String^ managedPortName = gcnew String(displayName.c_str());

        availablePorts->Add(managedPortName);
        comboBoxPorts->Items->Add(managedPortName);
    }

    if (availablePorts->Count == 0)
    {
        LogMessage(L"Не найдено ни одного COM-порта.");
    }
    else
    {
        LogMessage(L"Найдено портов: " + availablePorts->Count);
        comboBoxPorts->SelectedIndex = 0;
    }
}

void ModbusMasterForm::ConnectToPort()
{
    if (comboBoxPorts->SelectedIndex < 0 || availablePorts == nullptr)
    {
        MessageBox::Show(L"Выберите порт для подключения!", L"Ошибка", MessageBoxButtons::OK, MessageBoxIcon::Warning);
        return;
    }

    String^ selectedPort = availablePorts[comboBoxPorts->SelectedIndex];
    std::string selectedPortStd = msclr::interop::marshal_as<std::string>(selectedPort);

    LogMessage(L"Открытие порта " + selectedPort + L"...");

    hSerial = CreateFileA(selectedPortStd.c_str(), GENERIC_READ | GENERIC_WRITE, 0,
        NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);

    if (hSerial == INVALID_HANDLE_VALUE)
    {
        DWORD error = GetLastError();
        LogMessage(L"Ошибка! Не удалось открыть порт. Код ошибки: " + error);
        MessageBox::Show(L"Не удалось открыть порт!", L"Ошибка", MessageBoxButtons::OK, MessageBoxIcon::Error);
        return;
    }

    if (ConfigurePort())
    {
        ClosePort();
        LogMessage(L"Не удалось настроить порт.");
        return;
    }

    LogMessage(L"Порт успешно открыт и настроен!");

    buttonConnect->Enabled = false;
    buttonDisconnect->Enabled = true;
    comboBoxPorts->Enabled = false;
    buttonRefresh->Enabled = false;
    buttonScanDevices->Enabled = true;
    buttonOpenViewer->Enabled = true;

    labelStatus->Text = L"Статус: Подключен к " + comboBoxPorts->Text;
}

bool ModbusMasterForm::ConfigurePort()
{
    LogMessage(L"Настройка порта...");

    DCB dcbSerialParams = { 0 };
    dcbSerialParams.DCBlength = sizeof(dcbSerialParams);

    if (!GetCommState(hSerial, &dcbSerialParams))
    {
        LogMessage(L"Ошибка получения текущих параметров порта.");
        return true;
    }

    dcbSerialParams.BaudRate = baudRate;
    dcbSerialParams.ByteSize = byteSize;
    dcbSerialParams.StopBits = stopBits;
    dcbSerialParams.Parity = parity;

    dcbSerialParams.fOutxCtsFlow = FALSE;
    dcbSerialParams.fOutxDsrFlow = FALSE;
    dcbSerialParams.fDtrControl = DTR_CONTROL_DISABLE;
    dcbSerialParams.fRtsControl = RTS_CONTROL_DISABLE;

    if (!SetCommState(hSerial, &dcbSerialParams))
    {
        LogMessage(L"Ошибка установки параметров порта.");
        return true;
    }

    COMMTIMEOUTS timeouts = { 0 };
    timeouts.ReadIntervalTimeout = 50;
    timeouts.ReadTotalTimeoutConstant = 500;
    timeouts.ReadTotalTimeoutMultiplier = 10;
    timeouts.WriteTotalTimeoutConstant = 500;
    timeouts.WriteTotalTimeoutMultiplier = 10;

    if (!SetCommTimeouts(hSerial, &timeouts))
    {
        LogMessage(L"Ошибка установки таймаутов.");
        return true;
    }

    PurgeComm(hSerial, PURGE_RXCLEAR | PURGE_TXCLEAR);

    LogMessage(L"Порт сконфигурирован: 9600 бод, 8 бит данных, 1 стоп-бит, без контроля четности.");
    return false;
}

void ModbusMasterForm::DisconnectFromPort()
{
    ClosePort();

    buttonConnect->Enabled = true;
    buttonDisconnect->Enabled = false;
    comboBoxPorts->Enabled = true;
    buttonRefresh->Enabled = true;
    buttonScanDevices->Enabled = false;
    buttonOpenViewer->Enabled = false;

    labelStatus->Text = L"Статус: Отключен";
    LogMessage(L"Отключен от порта.");

    if (isViewerFormOpen)
    {
        for each (Form ^ form in Application::OpenForms)
        {
            RegistersViewerForm^ viewerForm = dynamic_cast<RegistersViewerForm^>(form);
            if (viewerForm != nullptr)
            {
                viewerForm->Close();
                break;
            }
        }
        isViewerFormOpen = false;
    }
}

void ModbusMasterForm::ClosePort()
{
    if (hSerial != INVALID_HANDLE_VALUE)
    {
        CloseHandle(hSerial);
        hSerial = INVALID_HANDLE_VALUE;
    }
}

uint16_t ModbusMasterForm::CalculateCRC16(const uint8_t* data, size_t length)
{
    unsigned char crcHi = 0xFF;
    unsigned char crcLo = 0xFF;
    unsigned int index;

    while (length--)
    {
        index = crcLo ^ *data++;
        crcLo = crcHi ^ auchCRCHi[index];
        crcHi = auchCRCLo[index];
    }

    return (crcHi << 8) | crcLo;
}

bool ModbusMasterForm::CheckCRC(const uint8_t* response, size_t length)
{
    if (length < 2) return false;
    uint16_t calculatedCRC = CalculateCRC16(response, length - 2);
    uint16_t responseCRC = (response[length - 1] << 8) | response[length - 2];
    return (calculatedCRC == responseCRC);
}

void ModbusMasterForm::ScanForDevices()
{
    if (hSerial == INVALID_HANDLE_VALUE)
    {
        MessageBox::Show(L"Порт не подключен!", L"Ошибка", MessageBoxButtons::OK, MessageBoxIcon::Warning);
        return;
    }

    LogMessage(L"Начинаю поиск устройств Modbus...");
    textBoxScanResult->Visible = true;
    labelFoundDevices->Visible = true;

    buttonScanDevices->Enabled = false;

    PurgeComm(hSerial, PURGE_RXCLEAR | PURGE_TXCLEAR);

    std::vector<int> foundDevices;
    std::wstringstream result;

    COMMTIMEOUTS originalTimeouts;
    GetCommTimeouts(hSerial, &originalTimeouts);

    COMMTIMEOUTS scanTimeouts;
    scanTimeouts.ReadIntervalTimeout = 50;
    scanTimeouts.ReadTotalTimeoutConstant = 100;
    scanTimeouts.ReadTotalTimeoutMultiplier = 0;
    scanTimeouts.WriteTotalTimeoutConstant = 100;
    scanTimeouts.WriteTotalTimeoutMultiplier = 0;

    SetCommTimeouts(hSerial, &scanTimeouts);

    std::vector<uint8_t> request;
    request.resize(8);

    for (int slaveID = MIN_SLAVE_ID; slaveID <= MAX_SLAVE_ID; slaveID++)
    {
        request[0] = slaveID;
        request[1] = 0x03;
        request[2] = 0x00;
        request[3] = 0x00;
        request[4] = 0x00;
        request[5] = 0x02;

        uint16_t crc = CalculateCRC16(request.data(), 6);
        request[6] = crc & 0xFF;
        request[7] = (crc >> 8) & 0xFF;

        PurgeComm(hSerial, PURGE_RXCLEAR);

        DWORD bytesWritten;
        if (!WriteFile(hSerial, request.data(), request.size(), &bytesWritten, NULL))
        {
            continue;
        }

        uint8_t response[256];
        DWORD totalBytesRead = 0;
        DWORD bytesRead = 0;
        DWORD startTime = GetTickCount64();

        while (GetTickCount64() - startTime < 200)
        {
            if (ReadFile(hSerial, response + totalBytesRead,
                sizeof(response) - totalBytesRead, &bytesRead, NULL))
            {
                if (bytesRead > 0)
                {
                    totalBytesRead += bytesRead;

                    if (totalBytesRead >= 5 && response[0] == slaveID)
                    {
                        if (CheckCRC(response, totalBytesRead))
                        {
                            foundDevices.push_back(slaveID);
                            LogMessage(L"Найдено устройство с ID: " + slaveID);

                            result << L"ID: " << slaveID << L"\r\n";
                            break;
                        }
                    }
                }
            }
            Sleep(10);
        }

        if (slaveID % 5 == 0)
        {
            labelStatus->Text = L"Сканирование... " + (slaveID * 100 / MAX_SLAVE_ID) + L"%";
            Application::DoEvents();
        }

        Sleep(50);
    }

    SetCommTimeouts(hSerial, &originalTimeouts);

    if (foundDevices.empty())
    {
        result << L"Устройства не найдены";
        LogMessage(L"Устройства не найдены");
    }
    else
    {
        LogMessage(L"Поиск завершен. Найдено устройств: " + foundDevices.size());
    }

    textBoxScanResult->Text = gcnew String(result.str().c_str());
    labelStatus->Text = L"Статус: Подключен к " + comboBoxPorts->Text;
    buttonScanDevices->Enabled = true;
}

void ModbusMasterForm::ScanAllRegisters()
{
    if (hSerial == INVALID_HANDLE_VALUE) return;

    const int TOTAL_32BIT_VALUES = 250;
    const int BATCH_SIZE = 20;

    uint8_t slaveID = static_cast<uint8_t>(numericSlaveID->Value);

    for (int startIdx = 0; startIdx < TOTAL_32BIT_VALUES; startIdx += BATCH_SIZE)
    {
        int currentBatchSize = min(BATCH_SIZE, TOTAL_32BIT_VALUES - startIdx);
        uint16_t startAddress = startIdx * 2;
        uint16_t registerCount = currentBatchSize * 2;

        std::vector<uint8_t> request;
        request.push_back(slaveID);
        request.push_back(0x03);
        request.push_back((startAddress >> 8) & 0xFF);
        request.push_back(startAddress & 0xFF);
        request.push_back((registerCount >> 8) & 0xFF);
        request.push_back(registerCount & 0xFF);

        uint16_t crc = CalculateCRC16(request.data(), request.size());
        request.push_back(crc & 0xFF);
        request.push_back((crc >> 8) & 0xFF);

        PurgeComm(hSerial, PURGE_RXCLEAR);

        DWORD bytesWritten;
        if (!WriteFile(hSerial, request.data(), request.size(), &bytesWritten, NULL))
        {
            LogMessage(L"Ошибка отправки запроса!");
            continue;
        }

        Sleep(100);

        uint8_t response[512];
        DWORD totalBytesRead = 0;
        DWORD bytesRead = 0;
        DWORD startTime = GetTickCount64();

        while ((GetTickCount64() - startTime) < 500)
        {
            if (ReadFile(hSerial, response + totalBytesRead,
                sizeof(response) - totalBytesRead, &bytesRead, NULL) && bytesRead > 0)
            {
                totalBytesRead += bytesRead;

                if (totalBytesRead >= 3 && response[1] == 0x03)
                {
                    int expectedLength = response[2] + 3;
                    if (totalBytesRead >= expectedLength)
                    {
                        break;
                    }
                }
            }
            Sleep(10);
        }

        if (totalBytesRead > 0 && CheckCRC(response, totalBytesRead))
        {
            uint8_t dataLength = response[2];
            int valueCount = dataLength / 4;

            for (int i = 0; i < valueCount; i++)
            {
                int globalIndex = startIdx + i;
                if (globalIndex >= TOTAL_32BIT_VALUES) break;

                uint16_t reg1 = (response[3 + i * 4] << 8) | response[3 + i * 4 + 1];
                uint16_t reg2 = (response[3 + i * 4 + 2] << 8) | response[3 + i * 4 + 3];

                uint16_t reg1_le = ((reg1 & 0xFF) << 8) | ((reg1 >> 8) & 0xFF);
                uint16_t reg2_le = ((reg2 & 0xFF) << 8) | ((reg2 >> 8) & 0xFF);
                uint32_t value32 = (reg2_le << 16) | reg1_le;

                float floatValue;
                memcpy(&floatValue, &value32, sizeof(float));
                int32_t intValue = static_cast<int32_t>(value32);

                allRegistersData->Add(Tuple::Create(globalIndex, floatValue, intValue));
            }
        }

        Sleep(50);
    }
}

void ModbusMasterForm::WriteRegisters(List<RegisterChange^>^ changes)
{
    if (hSerial == INVALID_HANDLE_VALUE)
    {
        MessageBox::Show(L"Порт не подключен!", L"Ошибка", MessageBoxButtons::OK, MessageBoxIcon::Warning);
        return;
    }

    LogMessage(String::Format(L"Начинаю запись {0} измененных регистров...", changes->Count));

    int successCount = 0;
    int failCount = 0;
    uint8_t slaveID = static_cast<uint8_t>(numericSlaveID->Value);

    for each (RegisterChange ^ change in changes)
    {
        uint16_t startAddress = change->RegisterIndex * 2;
        uint16_t registerCount = 2;

        std::vector<uint8_t> request;
        request.push_back(slaveID);
        request.push_back(0x10);
        request.push_back((startAddress >> 8) & 0xFF);
        request.push_back(startAddress & 0xFF);
        request.push_back((registerCount >> 8) & 0xFF);
        request.push_back(registerCount & 0xFF);
        request.push_back(registerCount * 2);

        uint16_t lowWord = change->NewValue & 0xFFFF;
        uint16_t highWord = (change->NewValue >> 16) & 0xFFFF;

        uint16_t lowWord_le = ((lowWord & 0xFF) << 8) | ((lowWord >> 8) & 0xFF);
        uint16_t highWord_le = ((highWord & 0xFF) << 8) | ((highWord >> 8) & 0xFF);

        request.push_back((lowWord_le >> 8) & 0xFF);
        request.push_back(lowWord_le & 0xFF);
        request.push_back((highWord_le >> 8) & 0xFF);
        request.push_back(highWord_le & 0xFF);

        uint16_t crc = CalculateCRC16(request.data(), request.size());
        request.push_back(crc & 0xFF);
        request.push_back((crc >> 8) & 0xFF);

        DWORD bytesWritten;
        if (!WriteFile(hSerial, request.data(), request.size(), &bytesWritten, NULL))
        {
            LogMessage(String::Format(L"Ошибка отправки запроса для регистра {0}", change->RegisterIndex));
            failCount++;
            continue;
        }

        Sleep(100);

        uint8_t response[256];
        DWORD totalBytesRead = 0;
        DWORD bytesRead = 0;
        DWORD startTime = GetTickCount64();

        while ((GetTickCount64() - startTime) < 500)
        {
            if (ReadFile(hSerial, response + totalBytesRead,
                sizeof(response) - totalBytesRead, &bytesRead, NULL) && bytesRead > 0)
            {
                totalBytesRead += bytesRead;

                if (totalBytesRead >= 8 && response[1] == 0x10)
                {
                    break;
                }
            }
            Sleep(10);
        }

        if (totalBytesRead >= 8 && CheckCRC(response, totalBytesRead))
        {
            LogMessage(String::Format(L"Регистр {0} успешно записан (значение: {1})",
                change->RegisterIndex, change->NewFloatValue));
            successCount++;
        }
        else
        {
            LogMessage(String::Format(L"Ошибка записи регистра {0}", change->RegisterIndex));
            failCount++;
        }

        Sleep(50);
    }

    LogMessage(String::Format(L"Запись завершена. Успешно: {0}, Ошибок: {1}", successCount, failCount));

    if (failCount == 0 && successCount > 0)
    {
        MessageBox::Show(L"Все изменения успешно записаны в устройство!", L"Успех", MessageBoxButtons::OK, MessageBoxIcon::Information);
    }
    else if (failCount > 0)
    {
        MessageBox::Show(String::Format(L"Запись завершена с ошибками. Успешно: {0}, Ошибок: {1}",
            successCount, failCount), L"Предупреждение", MessageBoxButtons::OK, MessageBoxIcon::Warning);
    }
}

void ModbusMasterForm::LogMessage(System::String^ message)
{
    textBoxLog->AppendText(DateTime::Now.ToString(L"HH:mm:ss") + L" - " + message + L"\r\n");
    textBoxLog->SelectionStart = textBoxLog->TextLength;
    textBoxLog->ScrollToCaret();
}

void ModbusMasterForm::LogMessage(const std::wstring& message)
{
    LogMessage(gcnew String(message.c_str()));
}

System::String^ ModbusMasterForm::StringToWString(const std::string& str)
{
    return gcnew String(str.c_str());
}

// ============================================================================
// Точка входа
// ============================================================================

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
        MessageBox::Show("Ошибка: " + ex->Message, "Ошибка", MessageBoxButtons::OK, MessageBoxIcon::Error);
    }

    return 0;
}