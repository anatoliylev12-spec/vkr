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

// Заголовок для маршалинга в C++/CLI
#include <msclr\marshal_cppstd.h>

using namespace System;
using namespace System::ComponentModel;
using namespace System::Collections;
using namespace System::Windows::Forms;
using namespace System::Data;
using namespace System::Drawing;
using namespace System::Collections::Generic;

/// <summary>
/// Класс главной формы
/// </summary>
public ref class ModbusMasterForm : public System::Windows::Forms::Form
{
public:
    ModbusMasterForm(void);

protected:
    ~ModbusMasterForm();

private:
    // Элементы управления для порта
    System::Windows::Forms::ComboBox^ comboBoxPorts;
    System::Windows::Forms::Button^ buttonRefresh;
    System::Windows::Forms::Button^ buttonConnect;
    System::Windows::Forms::Button^ buttonDisconnect;
    System::Windows::Forms::Button^ buttonScanDevices;
    System::Windows::Forms::TextBox^ textBoxLog;
    System::Windows::Forms::Label^ labelStatus;
    System::Windows::Forms::GroupBox^ groupBoxPort;
    System::Windows::Forms::CheckBox^ checkBoxShowAll;
    System::Windows::Forms::Button^ buttonClearLog;
    System::Windows::Forms::NumericUpDown^ numericSlaveID;
    System::Windows::Forms::Label^ labelSlaveID;
    System::Windows::Forms::Label^ labelFoundDevices;
    System::Windows::Forms::TextBox^ textBoxScanResult;
    System::Windows::Forms::ProgressBar^ progressBar;

    // Элементы управления для регистров
    System::Windows::Forms::DataGridView^ dataGridViewRegisters;
    System::Windows::Forms::DataGridView^ dataGridViewParameters;
    System::Windows::Forms::Button^ buttonExport;
    System::Windows::Forms::Button^ buttonRefreshRegisters;
    System::Windows::Forms::Button^ buttonStartAutoUpdate;
    System::Windows::Forms::Button^ buttonStopAutoUpdate;
    System::Windows::Forms::Button^ buttonToggleRegisters;
    System::Windows::Forms::Button^ buttonSendChanges;
    System::Windows::Forms::Label^ labelAutoUpdateStatus;
    System::Windows::Forms::GroupBox^ groupBoxRegisters;
    System::Windows::Forms::GroupBox^ groupBoxParameters;
    System::Windows::Forms::SaveFileDialog^ saveFileDialog;

    // Приватные поля
    HANDLE hSerial;
    List<String^>^ availablePorts;

    // Фиксированные параметры порта для RS-485
    DWORD baudRate;
    BYTE byteSize;
    BYTE stopBits;
    BYTE parity;

    // Константы
    const int MIN_SLAVE_ID = 1;
    const int MAX_SLAVE_ID = 20;
    const int MAX_REGISTER_ADDRESS = 500;   // 500 16-битных регистров
    const int BATCH_SIZE = 20;             // Максимум регистров за запрос
    const int MAX_WAIT_TICKS = 10;           // 8 * 50мс = 400мс таймаут

    // Хранилище данных регистров (КЛЮЧ = физический адрес)
    Dictionary<int, Tuple<float, int32_t>^>^ allRegistersData;

    // Данные для контролируемых параметров (фиксированные)
    List<String^>^ paramNames;
    List<int>^ paramAddresses;
    List<String^>^ paramUnits;

    // Измененные регистры
    Dictionary<int, uint32_t>^ changedRegistersValues;
    Dictionary<int, float>^ changedRegistersFloats;

    // Таймеры
    System::Windows::Forms::Timer^ updateTimer;   // Для автообновления (4 сек)
    System::Windows::Forms::Timer^ pollTimer;     // Для конечного автомата (50 мс)

    bool isAutoUpdating;
    bool isRegistersPanelExpanded;
    DateTime lastUpdateTime;
    int originalFormHeight;
    int originalRegistersGroupHeight;

    // Флаги синхронизации
    bool isUpdatingTables;
    bool isPollingInProgress;
    System::Object^ portLock;
    DWORD pollStartTime;

    // Конечный автомат опроса
    enum class PollState { IDLE, SEND, WAIT, NEXT, COMPLETE };
    PollState currentState;
    int currentAddress;
    int currentBatchSize;
    int waitCounter;
    int totalRegistersRead;

    // Буферы (статическое выделение)
    array<unsigned char>^ readBuffer;
    array<unsigned char>^ writeBuffer;

    System::ComponentModel::Container^ components;

    void InitializeComponent();
    void InitializeCustomComponent();
    void InitializeParameters();
    void InitializeTimers();
    void RefreshPorts();
    void ConnectToPort();
    bool ConfigurePort();
    void DisconnectFromPort();
    void ClosePort();
    uint16_t CalculateCRC16(const uint8_t* data, size_t length);
    bool CheckCRC(const uint8_t* response, size_t length);
    void ScanForDevices();
    void RefreshRegistersData();
    void WriteRegisters();
    void UpdateRegistersData();
    void UpdateFullRegistersTable();
    void UpdateParametersTable();
    void CollapseRegistersPanel();
    void ExpandRegistersPanel();
    void UpdateProgress(int value);
    void LogMessage(System::String^ message);
    void LogMessage(const std::wstring& message);
    System::String^ StringToWString(const std::string& str);

    // Обработчики таймеров
    void OnUpdateTimerTick(Object^ sender, EventArgs^ e);
    void OnPollTimerTick(Object^ sender, EventArgs^ e);

    // Методы конечного автомата
    void SendRequest();
    void CheckResponse();
    void MoveToNext();
    void CompletePolling();

    void MarkRegisterAsChanged(int registerAddress, uint32_t newValue, float newFloatValue);
    bool IsRegisterChanged(int registerAddress);
    void ClearChangedRegisters();

    // Вспомогательные методы (быстрые преобразования)
    float BytesToFloat(array<unsigned char>^ bytes, int offset);
    void FloatToBytes(float value, array<unsigned char>^ buffer, int offset);
    uint16_t BytesToUInt16(array<unsigned char>^ bytes, int offset);
    void UInt16ToBytes(uint16_t value, array<unsigned char>^ buffer, int offset);

    // Обработчики событий
    void Form1_Load(System::Object^ sender, System::EventArgs^ e);
    void buttonRefresh_Click(System::Object^ sender, System::EventArgs^ e);
    void buttonConnect_Click(System::Object^ sender, System::EventArgs^ e);
    void buttonDisconnect_Click(System::Object^ sender, System::EventArgs^ e);
    void buttonScanDevices_Click(System::Object^ sender, System::EventArgs^ e);
    void buttonClearLog_Click(System::Object^ sender, System::EventArgs^ e);
    void buttonExport_Click(System::Object^ sender, EventArgs^ e);
    void buttonRefreshRegisters_Click(System::Object^ sender, EventArgs^ e);
    void buttonStartAutoUpdate_Click(System::Object^ sender, EventArgs^ e);
    void buttonStopAutoUpdate_Click(System::Object^ sender, EventArgs^ e);
    void buttonToggleRegisters_Click(System::Object^ sender, EventArgs^ e);
    void buttonSendChanges_Click(System::Object^ sender, EventArgs^ e);
    void dataGridViewRegisters_CellValueChanged(Object^ sender, DataGridViewCellEventArgs^ e);
    void dataGridViewParameters_CellDoubleClick(Object^ sender, DataGridViewCellEventArgs^ e);
};