#pragma once

// Неуправляемые заголовки Windows
#include <windows.h>
#include <vector>
#include <string>
#include <set>
#include <sstream>
#include <iomanip>
#include <cmath>
#include <fcntl.h>
#include <limits>
#include <iostream>

// Заголовок для маршалинга в C++/CLR
#include <msclr\marshal_cppstd.h>

// Управляемые пространства имен
using namespace System;
using namespace System::ComponentModel;
using namespace System::Collections;
using namespace System::Windows::Forms;
using namespace System::Data;
using namespace System::Drawing;
using namespace System::Collections::Generic;

// Таблица CRC16 для старшего байта (Modbus)
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

// Таблица CRC16 для младшего байта (Modbus)
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

/// <summary>
/// Класс главной формы
/// </summary>
public ref class ModbusMasterForm : public System::Windows::Forms::Form
{
public:
    ModbusMasterForm(void)
    {
        InitializeComponent();
        InitializeCustomComponent();
        hSerial = INVALID_HANDLE_VALUE;

        // Фиксированные настройки для RS-485
        baudRate = CBR_9600;      // 9600 бод
        byteSize = 8;              // 8 бит данных
        stopBits = ONESTOPBIT;     // 1 стоп-бит
        parity = NOPARITY;         // Без четности
    }

protected:

    ~ModbusMasterForm()
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

private:
    // Элементы управления
    System::Windows::Forms::ComboBox^ comboBoxPorts;
    System::Windows::Forms::Button^ buttonRefresh;
    System::Windows::Forms::Button^ buttonConnect;
    System::Windows::Forms::Button^ buttonDisconnect;
    System::Windows::Forms::Button^ buttonScanDevices;
    System::Windows::Forms::ComboBox^ comboBoxCommands;
    System::Windows::Forms::Button^ buttonSend;
    System::Windows::Forms::TextBox^ textBoxRequest;
    System::Windows::Forms::TextBox^ textBoxResponse;
    System::Windows::Forms::TextBox^ textBoxLog;
    System::Windows::Forms::Label^ labelStatus;
    System::Windows::Forms::GroupBox^ groupBoxPort;
    System::Windows::Forms::GroupBox^ groupBoxCommands;
    System::Windows::Forms::GroupBox^ groupBoxResponse;
    System::Windows::Forms::CheckBox^ checkBoxShowAll;
    System::Windows::Forms::ListBox^ listBoxDevices;
    System::Windows::Forms::Label^ labelFoundDevices;
    System::Windows::Forms::NumericUpDown^ numericSlaveID;
    System::Windows::Forms::Label^ labelSlaveID;
    System::Windows::Forms::NumericUpDown^ numericStartAddress;
    System::Windows::Forms::Label^ labelStartAddress;
    System::Windows::Forms::NumericUpDown^ numericQuantity;
    System::Windows::Forms::Label^ labelQuantity;
    System::Windows::Forms::TextBox^ textBoxWriteData;
    System::Windows::Forms::Label^ labelWriteData;
    System::Windows::Forms::Button^ buttonClearLog;
    System::Windows::Forms::TextBox^ textBoxScanResult;

    // Приватные поля
    HANDLE hSerial;
    List<String^>^ availablePorts;

    // Фиксированные параметры порта для RS-485
    DWORD baudRate;
    BYTE byteSize;
    BYTE stopBits;
    BYTE parity;

    // Константы для поиска устройств
    const int MIN_SLAVE_ID = 1;
    const int MAX_SLAVE_ID = 20;
    const int SCAN_TIMEOUT_MS = 100;

    /// <summary>
    /// Обязательная переменная конструктора.
    /// </summary>
    System::ComponentModel::Container^ components;


       System::Windows::Forms::CheckBox^ checkBoxFloatMode;

    /// <summary>
    /// Инициализация пользовательских компонентов
    /// </summary>
    void InitializeCustomComponent()
    {
        // Настройка ComboBox для команд
        comboBoxCommands->Items->Add("03 - Read Holding Registers (32-bit)");
        comboBoxCommands->Items->Add("04 - Read Input Registers (32-bit)");
        comboBoxCommands->Items->Add("10 - Write Multiple Registers (32-bit)");
        comboBoxCommands->SelectedIndex = 0;

        // Настройка NumericUpDown
        numericSlaveID->Minimum = MIN_SLAVE_ID;
        numericSlaveID->Maximum = MAX_SLAVE_ID;
        numericSlaveID->Value = 1;

        numericStartAddress->Minimum = 0;
        numericStartAddress->Maximum = 65535;
        numericStartAddress->Value = 0;
        numericStartAddress->Hexadecimal = true;

        // Для 32-битных значений количество - это количество 32-битных чисел
        numericQuantity->Minimum = 1;
        numericQuantity->Maximum = 62; // 125/2, так как 2 регистра на 32-битное значение
        numericQuantity->Value = 1;

        // Добавляем чекбокс для float
        checkBoxFloatMode = gcnew System::Windows::Forms::CheckBox();
        checkBoxFloatMode->Text = L"Интерпретировать как float (32-bit)";
        checkBoxFloatMode->Location = System::Drawing::Point(225, 70);
        checkBoxFloatMode->Size = System::Drawing::Size(250, 30);
        checkBoxFloatMode->Checked = false;
        this->groupBoxCommands->Controls->Add(checkBoxFloatMode);

        // Подсказка для данных записи
        textBoxWriteData->Text = L"Например: 1234.56 7890.12 (float) или 12345678 87654321 (int32)";
        textBoxWriteData->ForeColor = System::Drawing::Color::Gray;

        // Настройка поля для результатов сканирования
        textBoxScanResult->ReadOnly = true;
        textBoxScanResult->Multiline = true;
        textBoxScanResult->ScrollBars = ScrollBars::Vertical;
    }

#pragma region Windows Form Designer generated code
    /// <summary>
    /// Требуемый метод для поддержки конструктора — не изменяйте 
    /// содержимое этого метода с помощью редактора кода.
    /// </summary>
    void InitializeComponent(void)
    {
        this->comboBoxPorts = (gcnew System::Windows::Forms::ComboBox());
        this->buttonRefresh = (gcnew System::Windows::Forms::Button());
        this->buttonConnect = (gcnew System::Windows::Forms::Button());
        this->buttonDisconnect = (gcnew System::Windows::Forms::Button());
        this->buttonScanDevices = (gcnew System::Windows::Forms::Button());
        this->comboBoxCommands = (gcnew System::Windows::Forms::ComboBox());
        this->buttonSend = (gcnew System::Windows::Forms::Button());
        this->textBoxRequest = (gcnew System::Windows::Forms::TextBox());
        this->textBoxResponse = (gcnew System::Windows::Forms::TextBox());
        this->textBoxLog = (gcnew System::Windows::Forms::TextBox());
        this->labelStatus = (gcnew System::Windows::Forms::Label());
        this->groupBoxPort = (gcnew System::Windows::Forms::GroupBox());
        this->checkBoxShowAll = (gcnew System::Windows::Forms::CheckBox());
        this->groupBoxCommands = (gcnew System::Windows::Forms::GroupBox());
        this->labelWriteData = (gcnew System::Windows::Forms::Label());
        this->textBoxWriteData = (gcnew System::Windows::Forms::TextBox());
        this->labelQuantity = (gcnew System::Windows::Forms::Label());
        this->numericQuantity = (gcnew System::Windows::Forms::NumericUpDown());
        this->labelStartAddress = (gcnew System::Windows::Forms::Label());
        this->numericStartAddress = (gcnew System::Windows::Forms::NumericUpDown());
        this->labelSlaveID = (gcnew System::Windows::Forms::Label());
        this->numericSlaveID = (gcnew System::Windows::Forms::NumericUpDown());
        this->groupBoxResponse = (gcnew System::Windows::Forms::GroupBox());
        this->textBoxScanResult = (gcnew System::Windows::Forms::TextBox());
        this->labelFoundDevices = (gcnew System::Windows::Forms::Label());
        this->buttonClearLog = (gcnew System::Windows::Forms::Button());
        this->groupBoxPort->SuspendLayout();
        this->groupBoxCommands->SuspendLayout();
        (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->numericQuantity))->BeginInit();
        (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->numericStartAddress))->BeginInit();
        (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->numericSlaveID))->BeginInit();
        this->groupBoxResponse->SuspendLayout();
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
        this->buttonRefresh->Size = System::Drawing::Size(109, 30);
        this->buttonRefresh->TabIndex = 1;
        this->buttonRefresh->Text = L"Обновить";
        this->buttonRefresh->UseVisualStyleBackColor = true;
        this->buttonRefresh->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonRefresh_Click);
        // 
        // buttonConnect
        // 
        this->buttonConnect->Location = System::Drawing::Point(290, 30);
        this->buttonConnect->Name = L"buttonConnect";
        this->buttonConnect->Size = System::Drawing::Size(131, 30);
        this->buttonConnect->TabIndex = 2;
        this->buttonConnect->Text = L"Подключить";
        this->buttonConnect->UseVisualStyleBackColor = true;
        this->buttonConnect->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonConnect_Click);
        // 
        // buttonDisconnect
        // 
        this->buttonDisconnect->Enabled = false;
        this->buttonDisconnect->Location = System::Drawing::Point(427, 30);
        this->buttonDisconnect->Name = L"buttonDisconnect";
        this->buttonDisconnect->Size = System::Drawing::Size(115, 30);
        this->buttonDisconnect->TabIndex = 3;
        this->buttonDisconnect->Text = L"Отключить";
        this->buttonDisconnect->UseVisualStyleBackColor = true;
        this->buttonDisconnect->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonDisconnect_Click);
        // 
        // buttonScanDevices
        // 
        this->buttonScanDevices->Location = System::Drawing::Point(601, 30);
        this->buttonScanDevices->Name = L"buttonScanDevices";
        this->buttonScanDevices->Size = System::Drawing::Size(100, 30);
        this->buttonScanDevices->TabIndex = 4;
        this->buttonScanDevices->Text = L"Поиск устройств";
        this->buttonScanDevices->UseVisualStyleBackColor = true;
        this->buttonScanDevices->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonScanDevices_Click);
        // 
        // comboBoxCommands
        // 
        this->comboBoxCommands->DropDownStyle = System::Windows::Forms::ComboBoxStyle::DropDownList;
        this->comboBoxCommands->FormattingEnabled = true;
        this->comboBoxCommands->Location = System::Drawing::Point(15, 35);
        this->comboBoxCommands->Name = L"comboBoxCommands";
        this->comboBoxCommands->Size = System::Drawing::Size(200, 28);
        this->comboBoxCommands->TabIndex = 5;
        this->comboBoxCommands->SelectedIndexChanged += gcnew System::EventHandler(this, &ModbusMasterForm::comboBoxCommands_SelectedIndexChanged);
        // 
        // buttonSend
        // 
        this->buttonSend->Location = System::Drawing::Point(470, 95);
        this->buttonSend->Name = L"buttonSend";
        this->buttonSend->Size = System::Drawing::Size(109, 30);
        this->buttonSend->TabIndex = 6;
        this->buttonSend->Text = L"Отправить";
        this->buttonSend->UseVisualStyleBackColor = true;
        this->buttonSend->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonSend_Click);
        // 
        // textBoxRequest
        // 
        this->textBoxRequest->Location = System::Drawing::Point(19, 151);
        this->textBoxRequest->Multiline = true;
        this->textBoxRequest->Name = L"textBoxRequest";
        this->textBoxRequest->ReadOnly = true;
        this->textBoxRequest->Size = System::Drawing::Size(440, 55);
        this->textBoxRequest->TabIndex = 7;
        this->textBoxRequest->Text = L"Запрос будет отображен здесь";
        // 
        // textBoxResponse
        // 
        this->textBoxResponse->Location = System::Drawing::Point(15, 30);
        this->textBoxResponse->Multiline = true;
        this->textBoxResponse->Name = L"textBoxResponse";
        this->textBoxResponse->ReadOnly = true;
        this->textBoxResponse->ScrollBars = System::Windows::Forms::ScrollBars::Vertical;
        this->textBoxResponse->Size = System::Drawing::Size(555, 100);
        this->textBoxResponse->TabIndex = 8;
        // 
        // textBoxLog
        // 
        this->textBoxLog->Location = System::Drawing::Point(27, 339);
        this->textBoxLog->Multiline = true;
        this->textBoxLog->Name = L"textBoxLog";
        this->textBoxLog->ReadOnly = true;
        this->textBoxLog->ScrollBars = System::Windows::Forms::ScrollBars::Vertical;
        this->textBoxLog->Size = System::Drawing::Size(560, 100);
        this->textBoxLog->TabIndex = 9;
        // 
        // labelStatus
        // 
        this->labelStatus->AutoSize = true;
        this->labelStatus->Location = System::Drawing::Point(23, 538);
        this->labelStatus->Name = L"labelStatus";
        this->labelStatus->Size = System::Drawing::Size(148, 20);
        this->labelStatus->TabIndex = 10;
        this->labelStatus->Text = L"Статус: Отключен";
        // 
        // groupBoxPort
        // 
        this->groupBoxPort->Controls->Add(this->checkBoxShowAll);
        this->groupBoxPort->Controls->Add(this->comboBoxPorts);
        this->groupBoxPort->Controls->Add(this->buttonScanDevices);
        this->groupBoxPort->Controls->Add(this->buttonRefresh);
        this->groupBoxPort->Controls->Add(this->buttonDisconnect);
        this->groupBoxPort->Controls->Add(this->buttonConnect);
        this->groupBoxPort->Location = System::Drawing::Point(12, 12);
        this->groupBoxPort->Name = L"groupBoxPort";
        this->groupBoxPort->Size = System::Drawing::Size(733, 100);
        this->groupBoxPort->TabIndex = 11;
        this->groupBoxPort->TabStop = false;
        this->groupBoxPort->Text = L"Управление портом (RS-485, 9600 бод)";
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
        // groupBoxCommands
        // 
        this->groupBoxCommands->Controls->Add(this->labelWriteData);
        this->groupBoxCommands->Controls->Add(this->textBoxWriteData);
        this->groupBoxCommands->Controls->Add(this->labelQuantity);
        this->groupBoxCommands->Controls->Add(this->numericQuantity);
        this->groupBoxCommands->Controls->Add(this->labelStartAddress);
        this->groupBoxCommands->Controls->Add(this->numericStartAddress);
        this->groupBoxCommands->Controls->Add(this->labelSlaveID);
        this->groupBoxCommands->Controls->Add(this->numericSlaveID);
        this->groupBoxCommands->Controls->Add(this->comboBoxCommands);
        this->groupBoxCommands->Controls->Add(this->buttonSend);
        this->groupBoxCommands->Controls->Add(this->textBoxRequest);
        this->groupBoxCommands->Location = System::Drawing::Point(603, 119);
        this->groupBoxCommands->Name = L"groupBoxCommands";
        this->groupBoxCommands->Size = System::Drawing::Size(585, 439);
        this->groupBoxCommands->TabIndex = 12;
        this->groupBoxCommands->TabStop = false;
        this->groupBoxCommands->Text = L"Команды Modbus";
        this->groupBoxCommands->Enter += gcnew System::EventHandler(this, &ModbusMasterForm::groupBoxCommands_Enter);
        // 
        // labelWriteData
        // 
        this->labelWriteData->AutoSize = true;
        this->labelWriteData->Location = System::Drawing::Point(225, 22);
        this->labelWriteData->Name = L"labelWriteData";
        this->labelWriteData->Size = System::Drawing::Size(72, 20);
        this->labelWriteData->TabIndex = 14;
        this->labelWriteData->Text = L"Данные:";
        this->labelWriteData->Visible = false;
        this->labelWriteData->Click += gcnew System::EventHandler(this, &ModbusMasterForm::labelWriteData_Click);
        // 
        // textBoxWriteData
        // 
        this->textBoxWriteData->Location = System::Drawing::Point(229, 99);
        this->textBoxWriteData->Name = L"textBoxWriteData";
        this->textBoxWriteData->Size = System::Drawing::Size(230, 26);
        this->textBoxWriteData->TabIndex = 13;
        this->textBoxWriteData->Visible = false;
        this->textBoxWriteData->Enter += gcnew System::EventHandler(this, &ModbusMasterForm::textBoxWriteData_Enter);
        this->textBoxWriteData->Leave += gcnew System::EventHandler(this, &ModbusMasterForm::textBoxWriteData_Leave);
        // 
        // labelQuantity
        // 
        this->labelQuantity->AutoSize = true;
        this->labelQuantity->Location = System::Drawing::Point(470, 10);
        this->labelQuantity->Name = L"labelQuantity";
        this->labelQuantity->Size = System::Drawing::Size(104, 20);
        this->labelQuantity->TabIndex = 12;
        this->labelQuantity->Text = L"Количество:";
        // 
        // numericQuantity
        // 
        this->numericQuantity->Location = System::Drawing::Point(470, 35);
        this->numericQuantity->Name = L"numericQuantity";
        this->numericQuantity->Size = System::Drawing::Size(100, 26);
        this->numericQuantity->TabIndex = 11;
        // 
        // labelStartAddress
        // 
        this->labelStartAddress->AutoSize = true;
        this->labelStartAddress->Location = System::Drawing::Point(121, 76);
        this->labelStartAddress->Name = L"labelStartAddress";
        this->labelStartAddress->Size = System::Drawing::Size(100, 20);
        this->labelStartAddress->TabIndex = 10;
        this->labelStartAddress->Text = L"Адрес (hex):";
        // 
        // numericStartAddress
        // 
        this->numericStartAddress->Hexadecimal = true;
        this->numericStartAddress->Location = System::Drawing::Point(125, 99);
        this->numericStartAddress->Maximum = System::Decimal(gcnew cli::array< System::Int32 >(4) { 65535, 0, 0, 0 });
        this->numericStartAddress->Name = L"numericStartAddress";
        this->numericStartAddress->Size = System::Drawing::Size(100, 26);
        this->numericStartAddress->TabIndex = 9;
        // 
        // labelSlaveID
        // 
        this->labelSlaveID->AutoSize = true;
        this->labelSlaveID->Location = System::Drawing::Point(6, 76);
        this->labelSlaveID->Name = L"labelSlaveID";
        this->labelSlaveID->Size = System::Drawing::Size(120, 20);
        this->labelSlaveID->TabIndex = 8;
        this->labelSlaveID->Text = L"ID устройства:";
        // 
        // numericSlaveID
        // 
        this->numericSlaveID->Location = System::Drawing::Point(10, 100);
        this->numericSlaveID->Maximum = System::Decimal(gcnew cli::array< System::Int32 >(4) { 247, 0, 0, 0 });
        this->numericSlaveID->Minimum = System::Decimal(gcnew cli::array< System::Int32 >(4) { 1, 0, 0, 0 });
        this->numericSlaveID->Name = L"numericSlaveID";
        this->numericSlaveID->Size = System::Drawing::Size(100, 26);
        this->numericSlaveID->TabIndex = 7;
        this->numericSlaveID->Value = System::Decimal(gcnew cli::array< System::Int32 >(4) { 1, 0, 0, 0 });
        // 
        // groupBoxResponse
        // 
        this->groupBoxResponse->Controls->Add(this->textBoxResponse);
        this->groupBoxResponse->Location = System::Drawing::Point(12, 193);
        this->groupBoxResponse->Name = L"groupBoxResponse";
        this->groupBoxResponse->Size = System::Drawing::Size(585, 140);
        this->groupBoxResponse->TabIndex = 13;
        this->groupBoxResponse->TabStop = false;
        this->groupBoxResponse->Text = L"Ответ устройства";
        // 
        // textBoxScanResult
        // 
        this->textBoxScanResult->Location = System::Drawing::Point(12, 153);
        this->textBoxScanResult->Name = L"textBoxScanResult";
        this->textBoxScanResult->Size = System::Drawing::Size(585, 26);
        this->textBoxScanResult->TabIndex = 14;
        this->textBoxScanResult->Visible = false;
        // 
        // labelFoundDevices
        // 
        this->labelFoundDevices->AutoSize = true;
        this->labelFoundDevices->Location = System::Drawing::Point(12, 133);
        this->labelFoundDevices->Name = L"labelFoundDevices";
        this->labelFoundDevices->Size = System::Drawing::Size(191, 20);
        this->labelFoundDevices->TabIndex = 15;
        this->labelFoundDevices->Text = L"Найденные устройства:";
        this->labelFoundDevices->Visible = false;
        // 
        // buttonClearLog
        // 
        this->buttonClearLog->Location = System::Drawing::Point(511, 538);
        this->buttonClearLog->Name = L"buttonClearLog";
        this->buttonClearLog->Size = System::Drawing::Size(72, 30);
        this->buttonClearLog->TabIndex = 16;
        this->buttonClearLog->Text = L"Очистить лог";
        this->buttonClearLog->UseVisualStyleBackColor = true;
        this->buttonClearLog->Click += gcnew System::EventHandler(this, &ModbusMasterForm::buttonClearLog_Click);
        // 
        // ModbusMasterForm
        // 
        this->AutoScaleDimensions = System::Drawing::SizeF(9, 20);
        this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
        this->ClientSize = System::Drawing::Size(1270, 705);
        this->Controls->Add(this->buttonClearLog);
        this->Controls->Add(this->labelFoundDevices);
        this->Controls->Add(this->textBoxScanResult);
        this->Controls->Add(this->groupBoxResponse);
        this->Controls->Add(this->groupBoxCommands);
        this->Controls->Add(this->groupBoxPort);
        this->Controls->Add(this->labelStatus);
        this->Controls->Add(this->textBoxLog);
        this->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
            static_cast<System::Byte>(204)));
        this->Margin = System::Windows::Forms::Padding(4, 5, 4, 5);
        this->Name = L"ModbusMasterForm";
        this->Text = L"Modbus Master для RS-485 устройства";
        this->Load += gcnew System::EventHandler(this, &ModbusMasterForm::Form1_Load);
        this->groupBoxPort->ResumeLayout(false);
        this->groupBoxPort->PerformLayout();
        this->groupBoxCommands->ResumeLayout(false);
        this->groupBoxCommands->PerformLayout();
        (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->numericQuantity))->EndInit();
        (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->numericStartAddress))->EndInit();
        (cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->numericSlaveID))->EndInit();
        this->groupBoxResponse->ResumeLayout(false);
        this->groupBoxResponse->PerformLayout();
        this->ResumeLayout(false);
        this->PerformLayout();

    }
#pragma endregion

    // Обработчики событий
    void Form1_Load(System::Object^ sender, System::EventArgs^ e)
    {
        RefreshPorts();
        LogMessage(L"Программа запущена. Режим RS-485, 9600 бод.");
    }

    void buttonRefresh_Click(System::Object^ sender, System::EventArgs^ e)
    {
        RefreshPorts();
    }

    void buttonConnect_Click(System::Object^ sender, System::EventArgs^ e)
    {
        ConnectToPort();
    }

    void buttonDisconnect_Click(System::Object^ sender, System::EventArgs^ e)
    {
        DisconnectFromPort();
    }

    void buttonScanDevices_Click(System::Object^ sender, System::EventArgs^ e)
    {
        ScanForDevices();
    }

    void buttonSend_Click(System::Object^ sender, System::EventArgs^ e)
    {
        SendModbusCommand();
    }

    void buttonClearLog_Click(System::Object^ sender, System::EventArgs^ e)
    {
        textBoxLog->Clear();
    }

    void comboBoxCommands_SelectedIndexChanged(System::Object^ sender, System::EventArgs^ e)
    {
        // Показываем поле для данных только для команды 10 (Write)
        bool isWriteCommand = (comboBoxCommands->SelectedIndex == 2);
        labelWriteData->Visible = isWriteCommand;
        textBoxWriteData->Visible = isWriteCommand;
    }

    void textBoxWriteData_Enter(System::Object^ sender, System::EventArgs^ e)
    {
        if (textBoxWriteData->Text == L"Например: 1234 5678 (hex)")
        {
            textBoxWriteData->Text = L"";
            textBoxWriteData->ForeColor = System::Drawing::Color::Black;
        }
    }

    void textBoxWriteData_Leave(System::Object^ sender, System::EventArgs^ e)
    {
        if (textBoxWriteData->Text->Trim() == L"")
        {
            textBoxWriteData->Text = L"Например: 1234 5678 (hex)";
            textBoxWriteData->ForeColor = System::Drawing::Color::Gray;
        }
    }

    // Функции работы с COM-портами
    void RefreshPorts()
    {
        comboBoxPorts->Items->Clear();

        // Создаем новый список или очищаем существующий
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

        // Копируем результаты в список и ComboBox
        for (const auto& port : portSet)
        {
            // Убираем префикс \\.\ для отображения
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


    void ConnectToPort()
    {
        if (comboBoxPorts->SelectedIndex < 0 || availablePorts == nullptr)
        {
            MessageBox::Show(L"Выберите порт для подключения!", L"Ошибка",
                MessageBoxButtons::OK, MessageBoxIcon::Warning);
            return;
        }

        // Получаем выбранный порт из управляемого списка
        String^ selectedPort = availablePorts[comboBoxPorts->SelectedIndex];

        // Конвертируем String^ в std::string для WinAPI
        std::string selectedPortStd = msclr::interop::marshal_as<std::string>(selectedPort);

        LogMessage(L"Открытие порта " + selectedPort + L"...");

        hSerial = CreateFileA(selectedPortStd.c_str(), GENERIC_READ | GENERIC_WRITE, 0,
            NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);

        if (hSerial == INVALID_HANDLE_VALUE)
        {
            DWORD error = GetLastError();
            LogMessage(L"Ошибка! Не удалось открыть порт. Код ошибки: " + error);
            MessageBox::Show(L"Не удалось открыть порт!", L"Ошибка",
                MessageBoxButtons::OK, MessageBoxIcon::Error);
            return;
        }

        if (ConfigurePort())
        {
            ClosePort();
            LogMessage(L"Не удалось настроить порт.");
            return;
        }

        LogMessage(L"Порт успешно открыт и настроен!");

        // Обновляем состояние кнопок
        buttonConnect->Enabled = false;
        buttonDisconnect->Enabled = true;
        comboBoxPorts->Enabled = false;
        buttonRefresh->Enabled = false;

        labelStatus->Text = L"Статус: Подключен к " + comboBoxPorts->Text;
    }

    bool ConfigurePort()
    {
        LogMessage(L"Настройка порта...");

        DCB dcbSerialParams = { 0 };
        dcbSerialParams.DCBlength = sizeof(dcbSerialParams);

        if (!GetCommState(hSerial, &dcbSerialParams))
        {
            LogMessage(L"Ошибка получения текущих параметров порта.");
            return true;
        }

        // Устанавливаем фиксированные параметры для RS-485
        dcbSerialParams.BaudRate = baudRate;
        dcbSerialParams.ByteSize = byteSize;
        dcbSerialParams.StopBits = stopBits;
        dcbSerialParams.Parity = parity;

        // Важно для RS-485: отключаем управление потоком
        dcbSerialParams.fOutxCtsFlow = FALSE;
        dcbSerialParams.fOutxDsrFlow = FALSE;
        dcbSerialParams.fDtrControl = DTR_CONTROL_DISABLE;
        dcbSerialParams.fRtsControl = RTS_CONTROL_DISABLE;

        if (!SetCommState(hSerial, &dcbSerialParams))
        {
            LogMessage(L"Ошибка установки параметров порта.");
            return true;
        }

        // Устанавливаем нормальные таймауты для работы
        COMMTIMEOUTS timeouts = { 0 };
        timeouts.ReadIntervalTimeout = 50;
        timeouts.ReadTotalTimeoutConstant = 500;  // Увеличиваем до 500 мс
        timeouts.ReadTotalTimeoutMultiplier = 10;
        timeouts.WriteTotalTimeoutConstant = 500; // Увеличиваем до 500 мс
        timeouts.WriteTotalTimeoutMultiplier = 10;

        if (!SetCommTimeouts(hSerial, &timeouts))
        {
            LogMessage(L"Ошибка установки таймаутов.");
            return true;
        }

        // Очищаем буферы
        PurgeComm(hSerial, PURGE_RXCLEAR | PURGE_TXCLEAR);

        LogMessage(L"Порт сконфигурирован: 9600 бод, 8 бит данных, 1 стоп-бит, без контроля четности.");
        return false;
    }

    void DisconnectFromPort()
    {
        ClosePort();

        buttonConnect->Enabled = true;
        buttonDisconnect->Enabled = false;
        comboBoxPorts->Enabled = true;
        buttonRefresh->Enabled = true;

        labelStatus->Text = L"Статус: Отключен";
        LogMessage(L"Отключен от порта.");
    }

    void ClosePort()
    {
        if (hSerial != INVALID_HANDLE_VALUE)
        {
            CloseHandle(hSerial);
            hSerial = INVALID_HANDLE_VALUE;
        }
    }

    // Функция вычисления CRC16 с использованием двух таблиц
    uint16_t CalculateCRC16(const uint8_t* data, size_t length)
    {
        unsigned char crcHi = 0xFF;  // Старший байт CRC
        unsigned char crcLo = 0xFF;  // Младший байт CRC
        unsigned int index;

        while (length--)
        {
            index = crcLo ^ *data++;
            crcLo = crcHi ^ auchCRCHi[index];
            crcHi = auchCRCLo[index];
        }

        // Возвращаем CRC в правильном порядке (младший байт, старший байт)
        return (crcHi << 8) | crcLo;
    }

    // Функция для проверки CRC полученного ответа
    bool CheckCRC(const uint8_t* response, size_t length)
    {
        if (length < 2) return false;

        // Вычисляем CRC для данных без последних двух байт
        uint16_t calculatedCRC = CalculateCRC16(response, length - 2);

        // Получаем CRC из ответа (в ответе сначала младший байт, потом старший)
        uint16_t responseCRC = (response[length - 1] << 8) | response[length - 2];

        return (calculatedCRC == responseCRC);
    }

    void SendModbusCommand()
    {
        if (hSerial == INVALID_HANDLE_VALUE)
        {
            MessageBox::Show(L"Порт не подключен!", L"Ошибка",
                MessageBoxButtons::OK, MessageBoxIcon::Warning);
            return;
        }

        uint8_t slaveID = (uint8_t)numericSlaveID->Value;
        uint16_t startAddress = (uint16_t)numericStartAddress->Value;
        uint16_t quantity32Bit = (uint16_t)numericQuantity->Value; // Количество 32-битных значений
        int commandIndex = comboBoxCommands->SelectedIndex;

        // Для Modbus запроса количество регистров = количество 32-битных значений * 2
        uint16_t registerCount = quantity32Bit * 2;

        // Проверяем максимальное количество регистров
        if (registerCount > 125)
        {
            MessageBox::Show(L"Превышено максимальное количество регистров (125)!",
                L"Ошибка", MessageBoxButtons::OK, MessageBoxIcon::Warning);
            return;
        }

        std::vector<uint8_t> request;
        request.push_back(slaveID);

        switch (commandIndex)
        {
        case 0: // 03 - Read Holding Registers (32-bit)
            request.push_back(0x03);
            request.push_back((startAddress >> 8) & 0xFF);
            request.push_back(startAddress & 0xFF);
            request.push_back((registerCount >> 8) & 0xFF);
            request.push_back(registerCount & 0xFF);
            break;

        case 1: // 04 - Read Input Registers (32-bit)
            request.push_back(0x04);
            request.push_back((startAddress >> 8) & 0xFF);
            request.push_back(startAddress & 0xFF);
            request.push_back((registerCount >> 8) & 0xFF);
            request.push_back(registerCount & 0xFF);
            break;

        case 2: // 10 - Write Multiple Registers (32-bit)
        {
            request.push_back(0x10);
            request.push_back((startAddress >> 8) & 0xFF);
            request.push_back(startAddress & 0xFF);
            request.push_back((registerCount >> 8) & 0xFF);
            request.push_back(registerCount & 0xFF);
            request.push_back(registerCount * 2); // Количество байт данных

            if (!ParseWriteData(request, quantity32Bit))
            {
                LogMessage(L"Ошибка парсинга данных для записи!");
                return;
            }
            break;
        }
        }

        // Вычисляем CRC16
        uint16_t crc = CalculateCRC16(request.data(), request.size());
        request.push_back(crc & 0xFF);       // Младший байт CRC
        request.push_back((crc >> 8) & 0xFF); // Старший байт CRC

        // Отображаем запрос
        DisplayRequest(request);

        // Очищаем буфер приема
        PurgeComm(hSerial, PURGE_RXCLEAR);

        // Отправляем запрос
        DWORD bytesWritten;
        BOOL success = WriteFile(hSerial, request.data(), request.size(), &bytesWritten, NULL);

        if (!success || bytesWritten != request.size())
        {
            LogMessage(L"Ошибка отправки запроса!");
            return;
        }

        LogMessage(L"Отправлено " + bytesWritten + L" байт (" + quantity32Bit + L" 32-битных значений)");

        Sleep(100);

        // Читаем ответ
        uint8_t response[256];
        DWORD totalBytesRead = 0;
        DWORD bytesRead = 0;
        DWORD startTime = GetTickCount();

        while ((GetTickCount() - startTime) < 500)
        {
            if (ReadFile(hSerial, response + totalBytesRead,
                sizeof(response) - totalBytesRead, &bytesRead, NULL) && bytesRead > 0)
            {
                totalBytesRead += bytesRead;

                if (totalBytesRead >= 5)
                {
                    int expectedLength = 0;
                    if (response[1] == 0x03 || response[1] == 0x04)
                    {
                        if (totalBytesRead >= 3)
                        {
                            expectedLength = response[2] + 3;
                        }
                    }
                    else if (response[1] == 0x10)
                    {
                        expectedLength = 8;
                    }

                    if (expectedLength > 0 && totalBytesRead >= expectedLength)
                    {
                        break;
                    }
                }
            }
            Sleep(10);
        }

        if (totalBytesRead > 0)
        {
            DisplayResponse(response, totalBytesRead);

            std::wstringstream logMsg;
            logMsg << L"Получен ответ (" << totalBytesRead << L" байт): ";
            for (DWORD i = 0; i < totalBytesRead; i++)
            {
                logMsg << std::hex << std::setw(2) << std::setfill(L'0')
                    << (int)response[i] << L" ";
            }
            LogMessage(logMsg.str());

            if (CheckCRC(response, totalBytesRead))
            {
                LogMessage(L"CRC проверка пройдена успешно!");

                if (response[1] == 0x03 || response[1] == 0x04)
                {
                    ParseAndDisplayRegisters(response, totalBytesRead);
                }
            }
            else
            {
                LogMessage(L"ОШИБКА: CRC проверка не пройдена!");
            }
        }
        else
        {
            LogMessage(L"Ответ не получен (таймаут)");
            textBoxResponse->Text = L"Ответ не получен. Проверьте подключение устройства.";
        }
    }



    void ParseAndDisplayRegisters(const uint8_t* response, DWORD length)
    {
        if (length < 3) return;

        uint8_t dataLength = response[2];
        int registerCount = dataLength / 2;
        int value32Count = registerCount / 2; // Количество 32-битных значений

        bool isFloatMode = checkBoxFloatMode->Checked;

        // Получаем начальный адрес и преобразуем в uint16_t
        uint16_t startAddress = static_cast<uint16_t>(numericStartAddress->Value);

        std::wstringstream ss;
        ss << L"\r\n\r\n=== 32-битные значения ===\r\n";

        if (isFloatMode)
            ss << L"Интерпретация: числа с плавающей точкой (float)\r\n";
        else
            ss << L"Интерпретация: 32-битные целые числа (int32)\r\n";

        ss << L"Формат: Modbus Word Swap (старший регистр, младший регистр)\r\n";
        ss << L"\r\n";

        for (int i = 0; i < value32Count; i++)
        {
            uint32_t regValue32;

            // Modbus word swap: [регистр n+1][регистр n]
            // Первый регистр содержит младшие 16 бит, второй - старшие 16 бит
            uint16_t lowWord = (response[3 + i * 4] << 8) | response[3 + i * 4 + 1];
            uint16_t highWord = (response[3 + i * 4 + 2] << 8) | response[3 + i * 4 + 3];

            regValue32 = (highWord << 16) | lowWord;

            // Вычисляем адреса регистров с правильным преобразованием типов
            uint16_t regAddr1 = startAddress + static_cast<uint16_t>(i * 2);
            uint16_t regAddr2 = startAddress + static_cast<uint16_t>(i * 2 + 1);

            ss << L"Значение " << i << L" [регистры 0x"
                << std::hex << std::setw(4) << std::setfill(L'0')
                << regAddr1 << L" и 0x"
                << std::hex << std::setw(4) << std::setfill(L'0')
                << regAddr2 << L"]: ";

            if (isFloatMode)
            {
                float floatValue;
                memcpy(&floatValue, &regValue32, sizeof(float));
                ss << floatValue;
                ss << L"\r\n   └─ Hex: 0x" << std::hex << std::setw(8) << std::setfill(L'0') << regValue32;
                ss << L" (int32: " << std::dec << static_cast<int32_t>(regValue32) << L")\r\n";
            }
            else
            {
                ss << L"0x" << std::hex << std::setw(8) << std::setfill(L'0') << regValue32;
                ss << L" (" << std::dec << static_cast<int32_t>(regValue32) << L")";
                ss << L"\r\n   └─ Как float: ";

                float floatValue;
                memcpy(&floatValue, &regValue32, sizeof(float));
                ss << floatValue << L"\r\n";
            }
        }

        // Добавляем сырые данные для отладки
        ss << L"\r\nСырые данные (" << dataLength << L" байт): ";
        for (DWORD i = 3; i < 3 + dataLength; i++)
        {
            ss << std::hex << std::setw(2) << std::setfill(L'0')
                << static_cast<int>(response[i]) << L" ";
        }

        textBoxResponse->Text += gcnew String(ss.str().c_str());
    }

    bool ParseWriteData(std::vector<uint8_t>& request, uint16_t quantity32Bit)
    {
        // quantity32Bit - количество 32-битных значений для записи
        String^ dataStr = textBoxWriteData->Text->Trim();
        bool isFloatMode = checkBoxFloatMode->Checked;

        if (dataStr == L"" || dataStr == L"Например: 1234.56 7890.12 (float) или 12345678 87654321 (int32)")
        {
            LogMessage(L"Введите данные для записи!");
            return false;
        }

        // Разбиваем строку по пробелам
        array<String^>^ values = dataStr->Split(' ');

        if (values->Length < quantity32Bit)
        {
            LogMessage(L"Ожидается " + quantity32Bit + L" значений, получено " + values->Length);
            return false;
        }

        for (int i = 0; i < quantity32Bit; i++)
        {
            try
            {
                uint32_t value32;

                if (isFloatMode)
                {
                    // Парсим как float
                    float floatValue = Convert::ToSingle(values[i]);
                    memcpy(&value32, &floatValue, sizeof(float));
                }
                else
                {
                    // Парсим как 32-битное целое (поддерживаем hex и dec)
                    String^ val = values[i];
                    if (val->StartsWith(L"0x") || val->StartsWith(L"0X"))
                    {
                        value32 = Convert::ToUInt32(val->Substring(2), 16);
                    }
                    else
                    {
                        value32 = Convert::ToUInt32(val);
                    }
                }

                // Формируем Modbus запрос с word swap (младший регистр, старший регистр)
                uint16_t lowWord = value32 & 0xFFFF;
                uint16_t highWord = (value32 >> 16) & 0xFFFF;

                // Записываем в порядке: сначала младший регистр, потом старший
                request.push_back((lowWord >> 8) & 0xFF);   // Старший байт младшего регистра
                request.push_back(lowWord & 0xFF);          // Младший байт младшего регистра
                request.push_back((highWord >> 8) & 0xFF);  // Старший байт старшего регистра
                request.push_back(highWord & 0xFF);         // Младший байт старшего регистра
            }
            catch (Exception^ ex)
            {
                LogMessage(L"Ошибка парсинга значения " + (i + 1) + L": " + values[i] + L" - " + ex->Message);
                return false;
            }
        }

        return true;
    }

    void DisplayRequest(const std::vector<uint8_t>& request)
    {
        std::wstringstream ss;

        int registerCount = 0;
        int value32Count = 0;

        if (request.size() >= 6)
        {
            registerCount = (request[4] << 8) | request[5];
            value32Count = registerCount / 2;
        }

        ss << L"Запрос (" << request.size() << L" байт, " << value32Count << L" 32-битных значений):\r\n";

        for (size_t i = 0; i < request.size(); i++)
        {
            ss << std::hex << std::setw(2) << std::setfill(L'0') << (int)request[i] << L" ";
            if ((i + 1) % 8 == 0 && i != request.size() - 1)
                ss << L"\r\n                ";
        }

        if (request.size() >= 2)
        {
            ss << L"\r\nCRC: "
                << std::hex << std::setw(2) << std::setfill(L'0') << (int)request[request.size() - 2]
                << L" "
                << std::hex << std::setw(2) << std::setfill(L'0') << (int)request[request.size() - 1];
        }

        textBoxRequest->Text = gcnew String(ss.str().c_str());
    }

    void DisplayResponse(const uint8_t* response, DWORD length)
    {
        std::wstringstream ss;

        for (DWORD i = 0; i < length; i++)
        {
            ss << std::hex << std::setw(2) << std::setfill(L'0') << (int)response[i] << L" ";
            // Каждые 8 байт добавляем перенос
            if ((i + 1) % 8 == 0 && i != length - 1)
                ss << L"\r\n";
        }

        // Добавляем информацию о CRC если есть
        if (length >= 2)
        {
            ss << L"\r\nCRC в ответе: "
                << std::hex << std::setw(2) << std::setfill(L'0') << (int)response[length - 2]
                << L" "
                << std::hex << std::setw(2) << std::setfill(L'0') << (int)response[length - 1];
        }

        textBoxResponse->Text = gcnew String(ss.str().c_str());

    }

    void ScanForDevices()
    {
        if (hSerial == INVALID_HANDLE_VALUE)
        {
            MessageBox::Show(L"Порт не подключен!", L"Ошибка",
                MessageBoxButtons::OK, MessageBoxIcon::Warning);
            return;
        }

        LogMessage(L"Начинаю поиск устройств Modbus (32-битный режим)...");
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
            request[5] = 0x02; // Читаем 2 регистра (одно 32-битное значение)

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

                                result << L"ID: " << slaveID << L" (ответ: ";
                                for (DWORD i = 0; i < totalBytesRead; i++)
                                    result << std::hex << std::setw(2) << std::setfill(L'0') << (int)response[i] << L" ";
                                result << L")\r\n";

                                break;
                            }
                        }
                    }
                }
                Sleep(10);
            }

            if (slaveID % 10 == 0)
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
    bool TestDevicePresence(uint8_t slaveID)
    {
        std::vector<uint8_t> request;
        request.push_back(slaveID);
        request.push_back(0x03); // Read Holding Registers
        request.push_back(0x00); // Старший байт адреса
        request.push_back(0x00); // Младший байт адреса
        request.push_back(0x00); // Старший байт количества
        request.push_back(0x01); // Младший байт количества

        uint16_t crc = CalculateCRC16(request.data(), request.size());
        request.push_back(crc & 0xFF);
        request.push_back((crc >> 8) & 0xFF);

        // Очищаем буфер
        PurgeComm(hSerial, PURGE_RXCLEAR);

        // Отправляем
        DWORD bytesWritten;
        if (!WriteFile(hSerial, request.data(), request.size(), &bytesWritten, NULL))
            return false;

        // Ждем переключения RS-485
        Sleep(50);

        // Читаем ответ
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
                        // Проверяем CRC
                        return CheckCRC(response, totalBytesRead);
                    }
                }
            }
            Sleep(50);
        }

        return false;
    }

    void LogMessage(System::String^ message)
    {
        textBoxLog->AppendText(DateTime::Now.ToString(L"HH:mm:ss") + L" - " + message + L"\r\n");
        // Прокручиваем вниз
        textBoxLog->SelectionStart = textBoxLog->TextLength;
        textBoxLog->ScrollToCaret();
    }


    void LogMessage(const std::wstring& message)
    {
        LogMessage(gcnew String(message.c_str()));
    }

    void LogMessage(int value)
    {
        LogMessage(value.ToString());
    }

    System::String^ StringToWString(const std::string& str)
    {
        return gcnew String(str.c_str());
    }
private: System::Void groupBoxCommands_Enter(System::Object^ sender, System::EventArgs^ e) {
}
private: System::Void labelWriteData_Click(System::Object^ sender, System::EventArgs^ e) {
}

private: System::Void button1_Click(System::Object^ sender, System::EventArgs^ e) {
}
};