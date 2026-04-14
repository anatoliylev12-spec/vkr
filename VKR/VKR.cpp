#include "VKR.h"

using namespace System;
using namespace System::Windows::Forms;
using namespace System::Xml;
using namespace System::Drawing;
using namespace System::Collections::Generic;
using namespace System::Threading;

// Таблицы CRC16
static unsigned char auchCRCHi[] = {
    0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81,
    0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0,
    0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01,
    0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41,
    0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81,
    0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0,
    0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01,
    0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40,
    0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81,
    0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0,
    0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01,
    0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41,
    0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81,
    0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0,
    0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01,
    0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41,
    0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81,
    0x40
};

static unsigned char auchCRCLo[] = {
    0x00, 0xC0, 0xC1, 0x01, 0xC3, 0x03, 0x02, 0xC2, 0xC6, 0x06, 0x07, 0xC7, 0x05, 0xC5, 0xC4,
    0x04, 0xCC, 0x0C, 0x0D, 0xCD, 0x0F, 0xCF, 0xCE, 0x0E, 0x0A, 0xCA, 0xCB, 0x0B, 0xC9, 0x09,
    0x08, 0xC8, 0xD8, 0x18, 0x19, 0xD9, 0x1B, 0xDB, 0xDA, 0x1A, 0x1E, 0xDE, 0xDF, 0x1F, 0xDD,
    0x1D, 0x1C, 0xDC, 0x14, 0xD4, 0xD5, 0x15, 0xD7, 0x17, 0x16, 0xD6, 0xD2, 0x12, 0x13, 0xD3,
    0x11, 0xD1, 0xD0, 0x10, 0xF0, 0x30, 0x31, 0xF1, 0x33, 0xF3, 0xF2, 0x32, 0x36, 0xF6, 0xF7,
    0x37, 0xF5, 0x35, 0x34, 0xF4, 0x3C, 0xFC, 0xFD, 0x3D, 0xFF, 0x3F, 0x3E, 0xFE, 0xFA, 0x3A,
    0x3B, 0xFB, 0x39, 0xF9, 0xF8, 0x38, 0x28, 0xE8, 0xE9, 0x29, 0xEB, 0x2B, 0x2A, 0xEA, 0xEE,
    0x2E, 0x2F, 0xEF, 0x2D, 0xED, 0xEC, 0x2C, 0xE4, 0x24, 0x25, 0xE5, 0x27, 0xE7, 0xE6, 0x26,
    0x22, 0xE2, 0xE3, 0x23, 0xE1, 0x21, 0x20, 0xE0, 0xA0, 0x60, 0x61, 0xA1, 0x63, 0xA3, 0xA2,
    0x62, 0x66, 0xA6, 0xA7, 0x67, 0xA5, 0x65, 0x64, 0xA4, 0x6C, 0xAC, 0xAD, 0x6D, 0xAF, 0x6F,
    0x6E, 0xAE, 0xAA, 0x6A, 0x6B, 0xAB, 0x69, 0xA9, 0xA8, 0x68, 0x78, 0xB8, 0xB9, 0x79, 0xBB,
    0x7B, 0x7A, 0xBA, 0xBE, 0x7E, 0x7F, 0xBF, 0x7D, 0xBD, 0xBC, 0x7C, 0xB4, 0x74, 0x75, 0xB5,
    0x77, 0xB7, 0xB6, 0x76, 0x72, 0xB2, 0xB3, 0x73, 0xB1, 0x71, 0x70, 0xB0, 0x50, 0x90, 0x91,
    0x51, 0x93, 0x53, 0x52, 0x92, 0x96, 0x56, 0x57, 0x97, 0x55, 0x95, 0x94, 0x54, 0x9C, 0x5C,
    0x5D, 0x9D, 0x5F, 0x9F, 0x9E, 0x5E, 0x5A, 0x9A, 0x9B, 0x5B, 0x99, 0x59, 0x58, 0x98, 0x88,
    0x48, 0x49, 0x89, 0x4B, 0x8B, 0x8A, 0x4A, 0x4E, 0x8E, 0x8F, 0x4F, 0x8D, 0x4D, 0x4C, 0x8C,
    0x44, 0x84, 0x85, 0x45, 0x87, 0x47, 0x46, 0x86, 0x82, 0x42, 0x43, 0x83, 0x41, 0x81, 0x80,
    0x40
};

// ============================================================================
// Конструктор и деструктор
// ============================================================================

ModbusMasterForm::ModbusMasterForm(void)
{
    InitializeComponent();
    InitializeCustomComponent();

    // Инициализация
    allRegistersData = gcnew Dictionary<int, Tuple<float, int32_t>^>();
    paramNames = gcnew List<String^>();
    paramAddresses = gcnew List<int>();
    paramUnits = gcnew List<String^>();
    changedRegistersValues = gcnew Dictionary<int, uint32_t>();
    changedRegistersFloats = gcnew Dictionary<int, float>();

    hSerial = INVALID_HANDLE_VALUE;
    baudRate = CBR_9600;
    byteSize = 8;
    stopBits = ONESTOPBIT;
    parity = NOPARITY;
    isAutoUpdating = false;
    isRegistersPanelExpanded = false;
    isUpdatingTables = false;
    isPollingInProgress = false;
    portLock = gcnew System::Object();
    pollStartTime = 0;

    // Буферы
    readBuffer = gcnew array<unsigned char>(2048);
    writeBuffer = gcnew array<unsigned char>(256);

    // Конечный автомат
    currentState = PollState::IDLE;
    currentAddress = 0;
    currentBatchSize = BATCH_SIZE;
    waitCounter = 0;
    totalRegistersRead = 0;

    InitializeParameters();
    InitializeTimers();

    originalFormHeight = this->Height;
    originalRegistersGroupHeight = groupBoxRegisters->Height;
    CollapseRegistersPanel();
}

ModbusMasterForm::~ModbusMasterForm()
{
    if (components) delete components;
    if (hSerial != INVALID_HANDLE_VALUE) ClosePort();
    if (updateTimer != nullptr) { updateTimer->Stop(); delete updateTimer; }
    if (pollTimer != nullptr) { pollTimer->Stop(); delete pollTimer; }
}

// ============================================================================
// Инициализация компонентов формы
// ============================================================================

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
    this->labelFoundDevices = (gcnew System::Windows::Forms::Label());
    this->textBoxScanResult = (gcnew System::Windows::Forms::TextBox());
    this->buttonClearLog = (gcnew System::Windows::Forms::Button());
    this->progressBar = (gcnew System::Windows::Forms::ProgressBar());
    this->dataGridViewRegisters = (gcnew System::Windows::Forms::DataGridView());
    this->dataGridViewParameters = (gcnew System::Windows::Forms::DataGridView());
    this->buttonExport = (gcnew System::Windows::Forms::Button());
    this->buttonRefreshRegisters = (gcnew System::Windows::Forms::Button());
    this->buttonStartAutoUpdate = (gcnew System::Windows::Forms::Button());
    this->buttonStopAutoUpdate = (gcnew System::Windows::Forms::Button());
    this->buttonToggleRegisters = (gcnew System::Windows::Forms::Button());
    this->buttonSendChanges = (gcnew System::Windows::Forms::Button());
    this->labelAutoUpdateStatus = (gcnew System::Windows::Forms::Label());
    this->groupBoxRegisters = (gcnew System::Windows::Forms::GroupBox());
    this->groupBoxParameters = (gcnew System::Windows::Forms::GroupBox());
    this->saveFileDialog = (gcnew System::Windows::Forms::SaveFileDialog());
    this->groupBoxPort->SuspendLayout();
    (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->numericSlaveID))->BeginInit();
    (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridViewRegisters))->BeginInit();
    (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridViewParameters))->BeginInit();
    this->groupBoxRegisters->SuspendLayout();
    this->groupBoxParameters->SuspendLayout();
    this->SuspendLayout();

    // comboBoxPorts
    this->comboBoxPorts->DropDownStyle = System::Windows::Forms::ComboBoxStyle::DropDownList;
    this->comboBoxPorts->FormattingEnabled = true;
    this->comboBoxPorts->Location = System::Drawing::Point(15, 30);
    this->comboBoxPorts->Name = L"comboBoxPorts";
    this->comboBoxPorts->Size = System::Drawing::Size(150, 28);
    this->comboBoxPorts->TabIndex = 0;

    // buttonRefresh
    this->buttonRefresh->Location = System::Drawing::Point(175, 30);
    this->buttonRefresh->Name = L"buttonRefresh";
    this->buttonRefresh->Size = System::Drawing::Size(100, 30);
    this->buttonRefresh->TabIndex = 1;
    this->buttonRefresh->Text = L"Обновить порты";
    this->buttonRefresh->UseVisualStyleBackColor = true;
    this->buttonRefresh->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonRefresh_Click);

    // buttonConnect
    this->buttonConnect->Location = System::Drawing::Point(285, 30);
    this->buttonConnect->Name = L"buttonConnect";
    this->buttonConnect->Size = System::Drawing::Size(100, 30);
    this->buttonConnect->TabIndex = 2;
    this->buttonConnect->Text = L"Подключить";
    this->buttonConnect->UseVisualStyleBackColor = true;
    this->buttonConnect->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonConnect_Click);

    // buttonDisconnect
    this->buttonDisconnect->Enabled = false;
    this->buttonDisconnect->Location = System::Drawing::Point(395, 30);
    this->buttonDisconnect->Name = L"buttonDisconnect";
    this->buttonDisconnect->Size = System::Drawing::Size(100, 30);
    this->buttonDisconnect->TabIndex = 3;
    this->buttonDisconnect->Text = L"Отключить";
    this->buttonDisconnect->UseVisualStyleBackColor = true;
    this->buttonDisconnect->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonDisconnect_Click);

    // buttonScanDevices
    this->buttonScanDevices->Location = System::Drawing::Point(15, 95);
    this->buttonScanDevices->Name = L"buttonScanDevices";
    this->buttonScanDevices->Size = System::Drawing::Size(150, 30);
    this->buttonScanDevices->TabIndex = 4;
    this->buttonScanDevices->Text = L"Поиск устройств";
    this->buttonScanDevices->UseVisualStyleBackColor = true;
    this->buttonScanDevices->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonScanDevices_Click);

    // textBoxLog
    this->textBoxLog->Location = System::Drawing::Point(12, 540);
    this->textBoxLog->Multiline = true;
    this->textBoxLog->Name = L"textBoxLog";
    this->textBoxLog->ReadOnly = true;
    this->textBoxLog->ScrollBars = System::Windows::Forms::ScrollBars::Vertical;
    this->textBoxLog->Size = System::Drawing::Size(1000, 140);
    this->textBoxLog->TabIndex = 5;

    // labelStatus
    this->labelStatus->AutoSize = true;
    this->labelStatus->Location = System::Drawing::Point(12, 695);
    this->labelStatus->Name = L"labelStatus";
    this->labelStatus->Size = System::Drawing::Size(148, 20);
    this->labelStatus->TabIndex = 6;
    this->labelStatus->Text = L"Статус: Отключен";

    // groupBoxPort
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
    this->groupBoxPort->Size = System::Drawing::Size(1000, 140);
    this->groupBoxPort->TabIndex = 7;
    this->groupBoxPort->TabStop = false;
    this->groupBoxPort->Text = L"Управление портом (RS-485, 9600 бод)";

    // labelSlaveID
    this->labelSlaveID->AutoSize = true;
    this->labelSlaveID->Location = System::Drawing::Point(505, 70);
    this->labelSlaveID->Name = L"labelSlaveID";
    this->labelSlaveID->Size = System::Drawing::Size(120, 20);
    this->labelSlaveID->TabIndex = 12;
    this->labelSlaveID->Text = L"ID устройства:";

    // numericSlaveID
    this->numericSlaveID->Location = System::Drawing::Point(505, 95);
    this->numericSlaveID->Maximum = System::Decimal(gcnew cli::array< System::Int32 >(4) { 247, 0, 0, 0 });
    this->numericSlaveID->Minimum = System::Decimal(gcnew cli::array< System::Int32 >(4) { 1, 0, 0, 0 });
    this->numericSlaveID->Name = L"numericSlaveID";
    this->numericSlaveID->Size = System::Drawing::Size(60, 26);
    this->numericSlaveID->TabIndex = 11;
    this->numericSlaveID->Value = System::Decimal(gcnew cli::array< System::Int32 >(4) { 1, 0, 0, 0 });

    // checkBoxShowAll
    this->checkBoxShowAll->AutoSize = true;
    this->checkBoxShowAll->Location = System::Drawing::Point(15, 65);
    this->checkBoxShowAll->Name = L"checkBoxShowAll";
    this->checkBoxShowAll->Size = System::Drawing::Size(240, 24);
    this->checkBoxShowAll->TabIndex = 5;
    this->checkBoxShowAll->Text = L"Показывать занятые порты";
    this->checkBoxShowAll->UseVisualStyleBackColor = true;

    // labelFoundDevices
    this->labelFoundDevices->AutoSize = true;
    this->labelFoundDevices->Location = System::Drawing::Point(12, 155);
    this->labelFoundDevices->Name = L"labelFoundDevices";
    this->labelFoundDevices->Size = System::Drawing::Size(191, 20);
    this->labelFoundDevices->TabIndex = 8;
    this->labelFoundDevices->Text = L"Найденные устройства:";
    this->labelFoundDevices->Visible = false;

    // textBoxScanResult
    this->textBoxScanResult->Location = System::Drawing::Point(12, 175);
    this->textBoxScanResult->Multiline = true;
    this->textBoxScanResult->Name = L"textBoxScanResult";
    this->textBoxScanResult->Size = System::Drawing::Size(1000, 60);
    this->textBoxScanResult->TabIndex = 9;
    this->textBoxScanResult->Visible = false;

    // buttonClearLog
    this->buttonClearLog->Location = System::Drawing::Point(940, 695);
    this->buttonClearLog->Name = L"buttonClearLog";
    this->buttonClearLog->Size = System::Drawing::Size(72, 30);
    this->buttonClearLog->TabIndex = 10;
    this->buttonClearLog->Text = L"Очистить лог";
    this->buttonClearLog->UseVisualStyleBackColor = true;
    this->buttonClearLog->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonClearLog_Click);

    // progressBar
    this->progressBar->Location = System::Drawing::Point(12, 250);
    this->progressBar->Name = L"progressBar";
    this->progressBar->Size = System::Drawing::Size(500, 23);
    this->progressBar->Minimum = 0;
    this->progressBar->Maximum = MAX_REGISTER_ADDRESS;
    this->progressBar->TabIndex = 11;

    // dataGridViewRegisters
    this->dataGridViewRegisters->AllowUserToAddRows = false;
    this->dataGridViewRegisters->AllowUserToDeleteRows = false;
    this->dataGridViewRegisters->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
    this->dataGridViewRegisters->Dock = System::Windows::Forms::DockStyle::Fill;
    this->dataGridViewRegisters->Location = System::Drawing::Point(3, 22);
    this->dataGridViewRegisters->Name = L"dataGridViewRegisters";
    this->dataGridViewRegisters->RowHeadersWidth = 62;
    this->dataGridViewRegisters->Size = System::Drawing::Size(600, 145);
    this->dataGridViewRegisters->TabIndex = 0;
    this->dataGridViewRegisters->CellValueChanged += gcnew System::Windows::Forms::DataGridViewCellEventHandler(this, &ModbusMasterForm::dataGridViewRegisters_CellValueChanged);

    // dataGridViewParameters
    this->dataGridViewParameters->AllowUserToAddRows = false;
    this->dataGridViewParameters->AllowUserToDeleteRows = false;
    this->dataGridViewParameters->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
    this->dataGridViewParameters->Dock = System::Windows::Forms::DockStyle::Fill;
    this->dataGridViewParameters->Location = System::Drawing::Point(3, 22);
    this->dataGridViewParameters->Name = L"dataGridViewParameters";
    this->dataGridViewParameters->ReadOnly = true;
    this->dataGridViewParameters->RowHeadersWidth = 62;
    this->dataGridViewParameters->Size = System::Drawing::Size(382, 145);
    this->dataGridViewParameters->TabIndex = 1;
    this->dataGridViewParameters->CellDoubleClick += gcnew System::Windows::Forms::DataGridViewCellEventHandler(this, &ModbusMasterForm::dataGridViewParameters_CellDoubleClick);

    // buttonExport
    this->buttonExport->Location = System::Drawing::Point(520, 250);
    this->buttonExport->Name = L"buttonExport";
    this->buttonExport->Size = System::Drawing::Size(100, 35);
    this->buttonExport->TabIndex = 12;
    this->buttonExport->Text = L"Экспорт в CSV";
    this->buttonExport->UseVisualStyleBackColor = true;
    this->buttonExport->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonExport_Click);

    // buttonRefreshRegisters
    this->buttonRefreshRegisters->Location = System::Drawing::Point(118, 290);
    this->buttonRefreshRegisters->Name = L"buttonRefreshRegisters";
    this->buttonRefreshRegisters->Size = System::Drawing::Size(100, 35);
    this->buttonRefreshRegisters->TabIndex = 2;
    this->buttonRefreshRegisters->Text = L"Обновить";
    this->buttonRefreshRegisters->UseVisualStyleBackColor = true;
    this->buttonRefreshRegisters->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonRefreshRegisters_Click);

    // buttonStartAutoUpdate
    this->buttonStartAutoUpdate->Location = System::Drawing::Point(224, 290);
    this->buttonStartAutoUpdate->Name = L"buttonStartAutoUpdate";
    this->buttonStartAutoUpdate->Size = System::Drawing::Size(120, 35);
    this->buttonStartAutoUpdate->TabIndex = 3;
    this->buttonStartAutoUpdate->Text = L"Автообновление ВКЛ";
    this->buttonStartAutoUpdate->UseVisualStyleBackColor = true;
    this->buttonStartAutoUpdate->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonStartAutoUpdate_Click);

    // buttonStopAutoUpdate
    this->buttonStopAutoUpdate->Enabled = false;
    this->buttonStopAutoUpdate->Location = System::Drawing::Point(350, 290);
    this->buttonStopAutoUpdate->Name = L"buttonStopAutoUpdate";
    this->buttonStopAutoUpdate->Size = System::Drawing::Size(120, 35);
    this->buttonStopAutoUpdate->TabIndex = 4;
    this->buttonStopAutoUpdate->Text = L"Автообновление ВЫКЛ";
    this->buttonStopAutoUpdate->UseVisualStyleBackColor = true;
    this->buttonStopAutoUpdate->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonStopAutoUpdate_Click);

    // buttonToggleRegisters
    this->buttonToggleRegisters->Location = System::Drawing::Point(476, 290);
    this->buttonToggleRegisters->Name = L"buttonToggleRegisters";
    this->buttonToggleRegisters->Size = System::Drawing::Size(140, 35);
    this->buttonToggleRegisters->TabIndex = 5;
    this->buttonToggleRegisters->Text = L"Развернуть таблицу регистров";
    this->buttonToggleRegisters->UseVisualStyleBackColor = true;
    this->buttonToggleRegisters->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonToggleRegisters_Click);

    // buttonSendChanges
    this->buttonSendChanges->Enabled = false;
    this->buttonSendChanges->Location = System::Drawing::Point(622, 290);
    this->buttonSendChanges->Name = L"buttonSendChanges";
    this->buttonSendChanges->Size = System::Drawing::Size(120, 35);
    this->buttonSendChanges->TabIndex = 6;
    this->buttonSendChanges->Text = L"Отправить изменения";
    this->buttonSendChanges->UseVisualStyleBackColor = true;
    this->buttonSendChanges->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonSendChanges_Click);

    // labelAutoUpdateStatus
    this->labelAutoUpdateStatus->AutoSize = true;
    this->labelAutoUpdateStatus->Location = System::Drawing::Point(12, 335);
    this->labelAutoUpdateStatus->Name = L"labelAutoUpdateStatus";
    this->labelAutoUpdateStatus->Size = System::Drawing::Size(232, 20);
    this->labelAutoUpdateStatus->TabIndex = 5;
    this->labelAutoUpdateStatus->Text = L"Автообновление: выключено";

    // groupBoxRegisters
    this->groupBoxRegisters->Controls->Add(this->dataGridViewRegisters);
    this->groupBoxRegisters->Location = System::Drawing::Point(12, 360);
    this->groupBoxRegisters->Name = L"groupBoxRegisters";
    this->groupBoxRegisters->Size = System::Drawing::Size(606, 170);
    this->groupBoxRegisters->TabIndex = 6;
    this->groupBoxRegisters->TabStop = false;
    this->groupBoxRegisters->Text = L"Все регистры [СВЕРНУТО]";

    // groupBoxParameters
    this->groupBoxParameters->Controls->Add(this->dataGridViewParameters);
    this->groupBoxParameters->Location = System::Drawing::Point(624, 360);
    this->groupBoxParameters->Name = L"groupBoxParameters";
    this->groupBoxParameters->Size = System::Drawing::Size(388, 170);
    this->groupBoxParameters->TabIndex = 7;
    this->groupBoxParameters->TabStop = false;
    this->groupBoxParameters->Text = L"Контролируемые параметры";

    // ModbusMasterForm
    this->AutoScaleDimensions = System::Drawing::SizeF(9, 20);
    this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
    this->ClientSize = System::Drawing::Size(1024, 740);
    this->Controls->Add(this->progressBar);
    this->Controls->Add(this->groupBoxParameters);
    this->Controls->Add(this->groupBoxRegisters);
    this->Controls->Add(this->labelAutoUpdateStatus);
    this->Controls->Add(this->buttonSendChanges);
    this->Controls->Add(this->buttonToggleRegisters);
    this->Controls->Add(this->buttonStopAutoUpdate);
    this->Controls->Add(this->buttonStartAutoUpdate);
    this->Controls->Add(this->buttonRefreshRegisters);
    this->Controls->Add(this->buttonExport);
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
    this->Text = L"Modbus Master для RS-485 устройства (Монитор регистров)";
    this->Load += gcnew System::EventHandler(this, &ModbusMasterForm::Form1_Load);

    this->groupBoxPort->ResumeLayout(false);
    this->groupBoxPort->PerformLayout();
    (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->numericSlaveID))->EndInit();
    (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridViewRegisters))->EndInit();
    (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridViewParameters))->EndInit();
    this->groupBoxRegisters->ResumeLayout(false);
    this->groupBoxParameters->ResumeLayout(false);
    this->ResumeLayout(false);
    this->PerformLayout();
}

void ModbusMasterForm::InitializeCustomComponent()
{
    numericSlaveID->Minimum = MIN_SLAVE_ID;
    numericSlaveID->Maximum = MAX_SLAVE_ID;
    numericSlaveID->Value = 1;
}

void ModbusMasterForm::InitializeParameters()
{
    paramNames->Clear();
    paramAddresses->Clear();
    paramUnits->Clear();

    paramNames->Add(L"Плотность");
    paramAddresses->Add(257);
    paramUnits->Add(L"кг/м³");

    paramNames->Add(L"Температура");
    paramAddresses->Add(259);
    paramUnits->Add(L"°C");

    paramNames->Add(L"Uout");
    paramAddresses->Add(264);
    paramUnits->Add(L"В");

    paramNames->Add(L"Период А");
    paramAddresses->Add(286);
    paramUnits->Add(L"мкс");

    paramNames->Add(L"Ku");
    paramAddresses->Add(285);
    paramUnits->Add(L"");

    paramNames->Add(L"Плотность NC");
    paramAddresses->Add(288);
    paramUnits->Add(L"кг/м³");

    paramNames->Add(L"Добротность");
    paramAddresses->Add(265);
    paramUnits->Add(L"");
}

void ModbusMasterForm::InitializeTimers()
{
    updateTimer = gcnew System::Windows::Forms::Timer();
    updateTimer->Interval = 4000;
    updateTimer->Tick += gcnew EventHandler(this, &ModbusMasterForm::OnUpdateTimerTick);

    pollTimer = gcnew System::Windows::Forms::Timer();
    pollTimer->Interval = 25;  // ← 25 мс вместо 50 (в 2 раза быстрее)
    pollTimer->Tick += gcnew EventHandler(this, &ModbusMasterForm::OnPollTimerTick);

    isAutoUpdating = false;
}

// ============================================================================
// Быстрые преобразования
// ============================================================================

union FloatUnion {
    float f;
    struct {
        unsigned char b0, b1, b2, b3;
    } bytes;
};

float ModbusMasterForm::BytesToFloat(array<unsigned char>^ bytes, int offset)
{
    FloatUnion fu;
    fu.bytes.b0 = bytes[offset + 1];
    fu.bytes.b1 = bytes[offset + 0];
    fu.bytes.b2 = bytes[offset + 3];
    fu.bytes.b3 = bytes[offset + 2];
    return fu.f;
}

void ModbusMasterForm::FloatToBytes(float value, array<unsigned char>^ buffer, int offset)
{
    FloatUnion fu;
    fu.f = value;
    buffer[offset + 0] = fu.bytes.b1;
    buffer[offset + 1] = fu.bytes.b0;
    buffer[offset + 2] = fu.bytes.b3;
    buffer[offset + 3] = fu.bytes.b2;
}

uint16_t ModbusMasterForm::BytesToUInt16(array<unsigned char>^ bytes, int offset)
{
    return (bytes[offset + 1] << 8) | bytes[offset];
}

void ModbusMasterForm::UInt16ToBytes(uint16_t value, array<unsigned char>^ buffer, int offset)
{
    buffer[offset] = value & 0xFF;
    buffer[offset + 1] = (value >> 8) & 0xFF;
}

// ============================================================================
// Работа с портом
// ============================================================================

void ModbusMasterForm::Form1_Load(System::Object^ sender, System::EventArgs^ e)
{
    RefreshPorts();
    LogMessage(L"Программа запущена. Режим RS-485, 9600 бод.");
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
                    }
                    else
                    {
                        HANDLE hTest = CreateFileA(portName.c_str(), GENERIC_READ | GENERIC_WRITE, 0,
                            NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);

                        if (hTest != INVALID_HANDLE_VALUE)
                        {
                            CloseHandle(hTest);
                            portSet.insert(portName);
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
    buttonRefreshRegisters->Enabled = true;
    buttonStartAutoUpdate->Enabled = true;

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

    // Короткие таймауты для быстрой работы
    COMMTIMEOUTS timeouts = { 0 };
    timeouts.ReadIntervalTimeout = 5;
    timeouts.ReadTotalTimeoutMultiplier = 0;
    timeouts.ReadTotalTimeoutConstant = 50;
    timeouts.WriteTotalTimeoutMultiplier = 0;
    timeouts.WriteTotalTimeoutConstant = 0;

    if (!SetCommTimeouts(hSerial, &timeouts))
    {
        LogMessage(L"Ошибка установки таймаутов.");
        return true;
    }

    SetupComm(hSerial, 4096, 4096);
    PurgeComm(hSerial, PURGE_RXCLEAR | PURGE_TXCLEAR);

    LogMessage(L"Порт сконфигурирован: 9600 бод, 8 бит, 1 стоп-бит");
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
    buttonRefreshRegisters->Enabled = false;
    buttonStartAutoUpdate->Enabled = false;
    buttonStopAutoUpdate->Enabled = false;

    if (isAutoUpdating)
    {
        isAutoUpdating = false;
        updateTimer->Stop();
        buttonStartAutoUpdate->Enabled = false;
        buttonStopAutoUpdate->Enabled = false;
        labelAutoUpdateStatus->Text = L"Автообновление: выключено";
    }

    currentState = PollState::IDLE;
    if (pollTimer) pollTimer->Stop();
    isPollingInProgress = false;

    labelStatus->Text = L"Статус: Отключен";
    LogMessage(L"Отключен от порта.");
}

void ModbusMasterForm::ClosePort()
{
    if (hSerial != INVALID_HANDLE_VALUE)
    {
        CloseHandle(hSerial);
        hSerial = INVALID_HANDLE_VALUE;
    }
}

// ============================================================================
// Modbus протокол
// ============================================================================

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

// ============================================================================
// Сканирование устройств
// ============================================================================

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

    std::vector<uint8_t> request;
    request.resize(8);

    for (int slaveID = MIN_SLAVE_ID; slaveID <= MAX_SLAVE_ID; slaveID++)
    {
        request[0] = slaveID;
        request[1] = 0x03;
        request[2] = 0x00;
        request[3] = 0x00;
        request[4] = 0x00;
        request[5] = 0x01;

        uint16_t crc = CalculateCRC16(request.data(), 6);
        request[6] = crc & 0xFF;
        request[7] = (crc >> 8) & 0xFF;

        PurgeComm(hSerial, PURGE_RXCLEAR);

        DWORD bytesWritten;
        if (!WriteFile(hSerial, request.data(), (DWORD)request.size(), &bytesWritten, NULL))
        {
            continue;
        }

        Sleep(100);

        uint8_t response[256];
        DWORD bytesRead = 0;
        ReadFile(hSerial, response, sizeof(response), &bytesRead, NULL);

        if (bytesRead >= 5 && response[0] == slaveID && CheckCRC(response, bytesRead))
        {
            foundDevices.push_back(slaveID);
            LogMessage(L"Найдено устройство с ID: " + slaveID);
            result << L"ID: " << slaveID << L"\r\n";
        }

        if (slaveID % 5 == 0)
        {
            labelStatus->Text = L"Сканирование... " + (slaveID * 100 / MAX_SLAVE_ID) + L"%";
            Application::DoEvents();
        }

        Sleep(50);
    }

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

// ============================================================================
// Конечный автомат опроса
// ============================================================================

void ModbusMasterForm::RefreshRegistersData()
{
    if (hSerial == INVALID_HANDLE_VALUE) return;
    if (isPollingInProgress)
    {
        LogMessage(L"Опрос уже выполняется");
        return;
    }

    isPollingInProgress = true;
    allRegistersData->Clear();
    currentState = PollState::SEND;
    currentAddress = 0;
    totalRegistersRead = 0;
    pollStartTime = GetTickCount();

    progressBar->Value = 0;
    progressBar->Maximum = MAX_REGISTER_ADDRESS;

    pollTimer->Start();
}

void ModbusMasterForm::OnPollTimerTick(Object^ sender, EventArgs^ e)
{
    if (Monitor::TryEnter(portLock))
    {
        try
        {
            switch (currentState)
            {
            case PollState::SEND:
                SendRequest();
                break;

            case PollState::WAIT:
                waitCounter++;
                if (waitCounter > MAX_WAIT_TICKS)
                {
                    LogMessage(String::Format(L"Таймаут для адреса {0}", currentAddress));
                    currentState = PollState::NEXT;
                }
                else
                {
                    CheckResponse();
                }
                break;

            case PollState::NEXT:
                MoveToNext();
                break;

            case PollState::COMPLETE:
                CompletePolling();
                break;

            default:
                break;
            }
        }
        finally
        {
            Monitor::Exit(portLock);
        }
    }
}

void ModbusMasterForm::SendRequest()
{
    if (currentAddress >= MAX_REGISTER_ADDRESS)
    {
        currentState = PollState::COMPLETE;
        return;
    }

    uint8_t slaveID = static_cast<uint8_t>(numericSlaveID->Value);
    currentBatchSize = min(BATCH_SIZE, MAX_REGISTER_ADDRESS - currentAddress);

    // Обновляем прогресс
    UpdateProgress(currentAddress);

    // Формируем запрос
    for (int i = 0; i < writeBuffer->Length; i++) writeBuffer[i] = 0;

    writeBuffer[0] = slaveID;
    writeBuffer[1] = 0x03;
    writeBuffer[2] = (currentAddress >> 8) & 0xFF;
    writeBuffer[3] = currentAddress & 0xFF;
    writeBuffer[4] = (currentBatchSize >> 8) & 0xFF;
    writeBuffer[5] = currentBatchSize & 0xFF;

    pin_ptr<unsigned char> pWriteBuffer = &writeBuffer[0];
    uint16_t crc = CalculateCRC16(pWriteBuffer, 6);
    writeBuffer[6] = crc & 0xFF;
    writeBuffer[7] = (crc >> 8) & 0xFF;

    // Отправка
    DWORD bytesWritten = 0;
    if (!WriteFile(hSerial, pWriteBuffer, 8, &bytesWritten, NULL))
    {
        LogMessage(String::Format(L"Ошибка отправки для адреса {0}", currentAddress));
        currentState = PollState::NEXT;
        return;
    }

    // Очищаем буфер чтения
    PurgeComm(hSerial, PURGE_RXCLEAR);

    currentState = PollState::WAIT;
    waitCounter = 0;
}

void ModbusMasterForm::CheckResponse()
{
    DWORD errors = 0;
    COMSTAT comStat = { 0 };
    ClearCommError(hSerial, &errors, &comStat);

    if (comStat.cbInQue == 0) return;

    // Ожидаемый размер ответа: 3 + (BATCH_SIZE * 2) + 2 = 3 + 40 + 2 = 45 байт
    int expectedSize = 3 + (currentBatchSize * 2) + 2;

    if (comStat.cbInQue < expectedSize) return;  // Ждем полный ответ

    // Читаем все данные сразу
    DWORD bytesRead = 0;
    pin_ptr<unsigned char> pReadBuffer = &readBuffer[0];

    if (!ReadFile(hSerial, pReadBuffer, expectedSize, &bytesRead, NULL))
    {
        currentState = PollState::NEXT;
        return;
    }

    if (bytesRead < expectedSize) return;

    // Проверяем CRC
    if (!CheckCRC(pReadBuffer, bytesRead))
    {
        currentState = PollState::NEXT;
        return;
    }

    if (readBuffer[1] == 0x03)
    {
        uint8_t dataLength = readBuffer[2];
        int registerCount = dataLength / 2;

        for (int i = 0; i < registerCount; i++)
        {
            int physicalAddress = currentAddress + i;
            if (physicalAddress >= MAX_REGISTER_ADDRESS) break;

            uint16_t value16 = BytesToUInt16(readBuffer, 3 + i * 2);

            allRegistersData[physicalAddress] = Tuple::Create(
                static_cast<float>(value16),
                static_cast<int32_t>(value16)
            );

            totalRegistersRead++;
        }
    }

    currentState = PollState::NEXT;
}

void ModbusMasterForm::MoveToNext()
{
    currentAddress += currentBatchSize;
    currentState = PollState::SEND;
}

void ModbusMasterForm::CompletePolling()
{
    pollTimer->Stop();
    isPollingInProgress = false;

    DWORD elapsed = (GetTickCount() - pollStartTime) / 1000;
    LogMessage(String::Format(L"Опрос завершен за {0} сек. Прочитано регистров: {1}",
        elapsed, totalRegistersRead));

    progressBar->Value = MAX_REGISTER_ADDRESS;

    this->BeginInvoke(gcnew Action(this, &ModbusMasterForm::UpdateRegistersData));
}

void ModbusMasterForm::UpdateProgress(int value)
{
    if (progressBar->InvokeRequired)
    {
        progressBar->BeginInvoke(gcnew Action<int>(this, &ModbusMasterForm::UpdateProgress), value);
        return;
    }
    progressBar->Value = value;
}

// ============================================================================
// Обновление таблиц (ПРАВИЛЬНОЕ ОТОБРАЖЕНИЕ - 1 регистр = 1 строка)
// ============================================================================

void ModbusMasterForm::UpdateRegistersData()
{
    UpdateFullRegistersTable();
    UpdateParametersTable();

    if (!isAutoUpdating)
    {
        ClearChangedRegisters();
    }
}

void ModbusMasterForm::UpdateFullRegistersTable()
{
    if (isUpdatingTables) return;
    isUpdatingTables = true;

    int firstDisplayedScrollRow = 0;
    if (dataGridViewRegisters->FirstDisplayedScrollingRowIndex >= 0)
        firstDisplayedScrollRow = dataGridViewRegisters->FirstDisplayedScrollingRowIndex;

    dataGridViewRegisters->Rows->Clear();
    dataGridViewRegisters->Columns->Clear();

    // Правильные колонки: адрес, HEX, DEC
    dataGridViewRegisters->Columns->Add(L"Address", L"Адрес");
    dataGridViewRegisters->Columns->Add(L"ValueHex", L"Значение (hex)");
    dataGridViewRegisters->Columns->Add(L"ValueDec", L"Значение (dec)");

    dataGridViewRegisters->Columns[0]->Width = 120;
    dataGridViewRegisters->Columns[1]->Width = 120;
    dataGridViewRegisters->Columns[2]->Width = 120;

    // Сортируем адреса
    List<int>^ addresses = gcnew List<int>(allRegistersData->Keys);
    addresses->Sort();

    for each (int address in addresses)
    {
        int value = allRegistersData[address]->Item2;

        // ПРАВИЛЬНО: один адрес = одна строка (не диапазон!)
        String^ addrStr = String::Format(L"{0} (0x{1:X4})", address, address);
        String^ hexStr = String::Format(L"0x{0:X4}", value & 0xFFFF);

        DataGridViewRow^ row = gcnew DataGridViewRow();
        row->CreateCells(dataGridViewRegisters);
        row->Cells[0]->Value = addrStr;
        row->Cells[1]->Value = hexStr;
        row->Cells[2]->Value = value.ToString();

        if (IsRegisterChanged(address))
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
    isUpdatingTables = false;
}

void ModbusMasterForm::UpdateParametersTable()
{
    if (isUpdatingTables) return;
    isUpdatingTables = true;

    dataGridViewParameters->Rows->Clear();
    dataGridViewParameters->Columns->Clear();
    dataGridViewParameters->Columns->Add(L"DateTime", L"Дата/Время");
    dataGridViewParameters->Columns->Add(L"Parameter", L"Параметр");
    dataGridViewParameters->Columns->Add(L"Address", L"Адрес");
    dataGridViewParameters->Columns->Add(L"Value", L"Значение");
    dataGridViewParameters->Columns->Add(L"Unit", L"Ед. изм.");

    dataGridViewParameters->Columns[0]->Width = 120;
    dataGridViewParameters->Columns[1]->Width = 100;
    dataGridViewParameters->Columns[2]->Width = 80;
    dataGridViewParameters->Columns[3]->Width = 100;
    dataGridViewParameters->Columns[4]->Width = 80;
    dataGridViewParameters->ReadOnly = true;

    String^ currentTime = DateTime::Now.ToString(L"dd.MM.yyyy HH:mm:ss");

    for (int i = 0; i < paramNames->Count; i++)
    {
        int value = 0;
        int physicalAddress = paramAddresses[i];

        if (allRegistersData != nullptr && allRegistersData->ContainsKey(physicalAddress))
        {
            value = allRegistersData[physicalAddress]->Item2;
        }

        String^ addrStr = String::Format(L"{0} (0x{1:X4})", physicalAddress, physicalAddress);
        String^ valueStr = value.ToString();

        int rowIndex = dataGridViewParameters->Rows->Add(
            currentTime, paramNames[i], addrStr, valueStr, paramUnits[i]);

        if (IsRegisterChanged(physicalAddress))
        {
            dataGridViewParameters->Rows[rowIndex]->DefaultCellStyle->BackColor = Color::LightYellow;
        }
    }

    dataGridViewParameters->AutoResizeColumns(DataGridViewAutoSizeColumnsMode::AllCells);
    isUpdatingTables = false;
}

void ModbusMasterForm::CollapseRegistersPanel()
{
    groupBoxRegisters->Height = 50;
    dataGridViewRegisters->Visible = false;
    groupBoxRegisters->Text = L"Все регистры [СВЕРНУТО]";
    buttonToggleRegisters->Text = L"Развернуть таблицу регистров";
    isRegistersPanelExpanded = false;
}

void ModbusMasterForm::ExpandRegistersPanel()
{
    groupBoxRegisters->Height = 475;
    dataGridViewRegisters->Visible = true;
    groupBoxRegisters->Text = L"Все регистры [РАЗВЕРНУТО]";
    buttonToggleRegisters->Text = L"Свернуть таблицу регистров";
    isRegistersPanelExpanded = true;
}

// ============================================================================
// Работа с изменениями регистров
// ============================================================================

void ModbusMasterForm::MarkRegisterAsChanged(int registerAddress, uint32_t newValue, float newFloatValue)
{
    changedRegistersValues[registerAddress] = newValue;
    changedRegistersFloats[registerAddress] = newFloatValue;
    buttonSendChanges->Enabled = true;
}

bool ModbusMasterForm::IsRegisterChanged(int registerAddress)
{
    return changedRegistersValues->ContainsKey(registerAddress);
}

void ModbusMasterForm::ClearChangedRegisters()
{
    changedRegistersValues->Clear();
    changedRegistersFloats->Clear();
    buttonSendChanges->Enabled = false;
}

// ============================================================================
// Запись регистров
// ============================================================================

void ModbusMasterForm::WriteRegisters()
{
    if (hSerial == INVALID_HANDLE_VALUE) return;

    Monitor::Enter(portLock);

    try
    {
        LogMessage(String::Format(L"Начинаю запись {0} измененных регистров...",
            changedRegistersValues->Count));

        int successCount = 0;
        int failCount = 0;
        uint8_t slaveID = static_cast<uint8_t>(numericSlaveID->Value);

        for each (auto kvp in changedRegistersValues)
        {
            int registerAddress = kvp.Key;
            uint32_t newValue = kvp.Value;
            float newFloatValue = changedRegistersFloats[registerAddress];

            array<unsigned char>^ request = gcnew array<unsigned char>(8);
            request[0] = slaveID;
            request[1] = 0x06;  // Write Single Register
            request[2] = (registerAddress >> 8) & 0xFF;
            request[3] = registerAddress & 0xFF;
            request[4] = (newValue >> 8) & 0xFF;
            request[5] = newValue & 0xFF;

            pin_ptr<unsigned char> pRequest = &request[0];
            uint16_t crc = CalculateCRC16(pRequest, 6);
            request[6] = crc & 0xFF;
            request[7] = (crc >> 8) & 0xFF;

            DWORD bytesWritten;
            if (!WriteFile(hSerial, pRequest, 8, &bytesWritten, NULL))
            {
                LogMessage(String::Format(L"Ошибка отправки для регистра {0}", registerAddress));
                failCount++;
                continue;
            }

            Sleep(100);

            uint8_t response[256];
            DWORD bytesRead = 0;
            ReadFile(hSerial, response, sizeof(response), &bytesRead, NULL);

            if (bytesRead >= 8 && CheckCRC(response, bytesRead) && response[1] == 0x06)
            {
                LogMessage(String::Format(L"Регистр {0} успешно записан (значение: {1})",
                    registerAddress, newFloatValue));
                successCount++;
            }
            else
            {
                LogMessage(String::Format(L"Ошибка записи регистра {0}", registerAddress));
                failCount++;
            }

            Sleep(50);
        }

        LogMessage(String::Format(L"Запись завершена. Успешно: {0}, Ошибок: {1}",
            successCount, failCount));
    }
    finally
    {
        Monitor::Exit(portLock);
    }
}

// ============================================================================
// Обработчики событий кнопок
// ============================================================================

void ModbusMasterForm::OnUpdateTimerTick(Object^ sender, EventArgs^ e)
{
    if (!isAutoUpdating) return;
    if (isPollingInProgress) return;

    lastUpdateTime = DateTime::Now;
    labelAutoUpdateStatus->Text = String::Format(
        L"Автообновление: активно (последнее: {0:HH:mm:ss})", lastUpdateTime);

    RefreshRegistersData();
}

void ModbusMasterForm::buttonRefresh_Click(System::Object^ sender, System::EventArgs^ e)
{
    RefreshPorts();
}

void ModbusMasterForm::buttonConnect_Click(System::Object^ sender, System::EventArgs^ e)
{
    ConnectToPort();
}

void ModbusMasterForm::buttonDisconnect_Click(System::Object^ sender, System::EventArgs^ e)
{
    DisconnectFromPort();
}

void ModbusMasterForm::buttonScanDevices_Click(System::Object^ sender, System::EventArgs^ e)
{
    ScanForDevices();
}

void ModbusMasterForm::buttonClearLog_Click(System::Object^ sender, System::EventArgs^ e)
{
    textBoxLog->Clear();
}

void ModbusMasterForm::buttonRefreshRegisters_Click(System::Object^ sender, EventArgs^ e)
{
    if (hSerial != INVALID_HANDLE_VALUE)
    {
        if (!isPollingInProgress)
        {
            RefreshRegistersData();
        }
        else
        {
            LogMessage(L"Опрос уже выполняется, подождите...");
        }
    }
    else
    {
        MessageBox::Show(L"Порт не подключен!", L"Ошибка", MessageBoxButtons::OK, MessageBoxIcon::Warning);
    }
}

void ModbusMasterForm::buttonStartAutoUpdate_Click(Object^ sender, EventArgs^ e)
{
    if (hSerial != INVALID_HANDLE_VALUE)
    {
        isAutoUpdating = true;
        updateTimer->Start();
        buttonStartAutoUpdate->Enabled = false;
        buttonStopAutoUpdate->Enabled = true;
        labelAutoUpdateStatus->Text = L"Автообновление: активно (обновление каждые 4 сек)";
        RefreshRegistersData();
    }
    else
    {
        MessageBox::Show(L"Порт не подключен!", L"Ошибка", MessageBoxButtons::OK, MessageBoxIcon::Warning);
    }
}

void ModbusMasterForm::buttonStopAutoUpdate_Click(Object^ sender, EventArgs^ e)
{
    isAutoUpdating = false;
    updateTimer->Stop();
    buttonStartAutoUpdate->Enabled = true;
    buttonStopAutoUpdate->Enabled = false;
    labelAutoUpdateStatus->Text = L"Автообновление: выключено";
}

void ModbusMasterForm::buttonToggleRegisters_Click(Object^ sender, EventArgs^ e)
{
    if (isRegistersPanelExpanded)
        CollapseRegistersPanel();
    else
        ExpandRegistersPanel();
}

void ModbusMasterForm::buttonSendChanges_Click(System::Object^ sender, EventArgs^ e)
{
    if (changedRegistersValues->Count == 0)
    {
        MessageBox::Show(L"Нет изменений для отправки!", L"Информация",
            MessageBoxButtons::OK, MessageBoxIcon::Information);
        return;
    }

    if (hSerial != INVALID_HANDLE_VALUE)
    {
        WriteRegisters();
        ClearChangedRegisters();
        RefreshRegistersData();
    }
    else
    {
        MessageBox::Show(L"Порт не подключен!", L"Ошибка",
            MessageBoxButtons::OK, MessageBoxIcon::Warning);
    }
}

void ModbusMasterForm::buttonExport_Click(Object^ sender, EventArgs^ e)
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

// ============================================================================
// Обработчики таблиц
// ============================================================================

void ModbusMasterForm::dataGridViewRegisters_CellValueChanged(Object^ sender, DataGridViewCellEventArgs^ e)
{
    if (e->RowIndex >= 0 && e->ColumnIndex == 2)
    {
        String^ addrStr = dataGridViewRegisters->Rows[e->RowIndex]->Cells[0]->Value->ToString();
        int spacePos = addrStr->IndexOf(L' ');
        if (spacePos > 0)
        {
            int physicalAddress = Int32::Parse(addrStr->Substring(0, spacePos));

            String^ newValueStr = dataGridViewRegisters->Rows[e->RowIndex]->Cells[2]->Value->ToString();

            try
            {
                int newValue = Convert::ToInt32(newValueStr);

                dataGridViewRegisters->Rows[e->RowIndex]->Cells[1]->Value = String::Format(L"0x{0:X4}", newValue & 0xFFFF);

                MarkRegisterAsChanged(physicalAddress, static_cast<uint32_t>(newValue), static_cast<float>(newValue));
                dataGridViewRegisters->Rows[e->RowIndex]->DefaultCellStyle->BackColor = Color::LightYellow;
                UpdateParametersTable();
            }
            catch (Exception^ ex)
            {
                MessageBox::Show(L"Ошибка ввода значения: " + ex->Message, L"Ошибка",
                    MessageBoxButtons::OK, MessageBoxIcon::Warning);
                UpdateFullRegistersTable();
            }
        }
    }
}

void ModbusMasterForm::dataGridViewParameters_CellDoubleClick(Object^ sender, DataGridViewCellEventArgs^ e)
{
    if (e->RowIndex >= 0 && e->RowIndex < paramNames->Count)
    {
        int physicalAddress = paramAddresses[e->RowIndex];

        for (int i = 0; i < dataGridViewRegisters->Rows->Count; i++)
        {
            DataGridViewRow^ row = dataGridViewRegisters->Rows[i];
            if (row->Cells[0]->Value != nullptr)
            {
                String^ addrStr = row->Cells[0]->Value->ToString();
                if (addrStr->Contains(String::Format(L"{0} (", physicalAddress)))
                {
                    dataGridViewRegisters->FirstDisplayedScrollingRowIndex = i;
                    dataGridViewRegisters->Rows[i]->Selected = true;
                    dataGridViewRegisters->CurrentCell = dataGridViewRegisters->Rows[i]->Cells[0];
                    break;
                }
            }
        }

        if (!isRegistersPanelExpanded)
        {
            ExpandRegistersPanel();
        }
    }
}

// ============================================================================
// Логирование
// ============================================================================

void ModbusMasterForm::LogMessage(System::String^ message)
{
    if (textBoxLog->InvokeRequired)
    {
        textBoxLog->BeginInvoke(gcnew Action<String^>(this, &ModbusMasterForm::LogMessage), message);
        return;
    }
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