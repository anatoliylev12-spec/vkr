// Table.h
#pragma once

#include <windows.h>
#include <vector>
#include <string>
#include <sstream>
#include <iomanip>
#include <cmath>

using namespace System;
using namespace System::Windows::Forms;
using namespace System::Drawing;
using namespace System::Collections::Generic;
using namespace System::Timers;

/// <summary>
/// Структура для хранения параметров
/// </summary>
public ref class ParameterInfo
{
public:
    String^ Name;
    int RegisterAddress;
    float Value;
    String^ Unit;

    ParameterInfo(String^ name, int address, String^ unit)
    {
        Name = name;
        RegisterAddress = address;
        Value = 0.0f;
        Unit = unit;
    }
};

/// <summary>
/// Структура для хранения изменений регистров
/// </summary>
public ref class RegisterChange
{
public:
    int RegisterIndex;  // 32-битный индекс
    uint32_t NewValue;
    float NewFloatValue;

    RegisterChange(int index, uint32_t value, float floatValue)
    {
        RegisterIndex = index;
        NewValue = value;
        NewFloatValue = floatValue;
    }
};

/// <summary>
/// Форма для отображения всех регистров в виде таблицы
/// </summary>
public ref class RegistersViewerForm : public System::Windows::Forms::Form
{
public:
    RegistersViewerForm(void);

    // Делегаты для обновления данных и записи регистров
    delegate void RefreshDataDelegate();
    delegate void WriteRegistersDelegate(List<RegisterChange^>^ changes);

    // Свойства для хранения функций обратного вызова
    property RefreshDataDelegate^ OnRefreshData;
    property WriteRegistersDelegate^ OnWriteRegisters;

    void UpdateRegistersData(List<System::Tuple<int, float, int32_t>^>^ registersData);

protected:
    ~RegistersViewerForm();

private:
    System::Windows::Forms::DataGridView^ dataGridViewRegisters;
    System::Windows::Forms::DataGridView^ dataGridViewParameters;
    System::Windows::Forms::Button^ buttonExport;
    System::Windows::Forms::Button^ buttonRefresh;
    System::Windows::Forms::Button^ buttonStartAutoUpdate;
    System::Windows::Forms::Button^ buttonStopAutoUpdate;
    System::Windows::Forms::Button^ buttonToggleRegisters;
    System::Windows::Forms::Button^ buttonSendChanges;
    System::Windows::Forms::Label^ labelAutoUpdateStatus;
    System::Windows::Forms::GroupBox^ groupBoxRegisters;
    System::Windows::Forms::GroupBox^ groupBoxParameters;
    System::Windows::Forms::SaveFileDialog^ saveFileDialog;
    System::ComponentModel::Container^ components;

    System::Timers::Timer^ updateTimer;
    bool isAutoUpdating;
    bool isRegistersPanelExpanded;

    List<ParameterInfo^>^ parametersList;
    List<System::Tuple<int, float, int32_t>^>^ allRegistersData;
    Dictionary<int, RegisterChange^>^ changedRegisters;
    DateTime lastUpdateTime;
    int originalFormHeight;
    int originalRegistersGroupHeight;

    void InitializeComponent();
    void InitializeParameters();
    void InitializeTimer();
    void UpdateFullRegistersTable(List<System::Tuple<int, float, int32_t>^>^ registersData);
    void UpdateParametersTable(List<System::Tuple<int, float, int32_t>^>^ registersData);
    void CollapseRegistersPanel();
    void ExpandRegistersPanel();

    void dataGridViewRegisters_CellValueChanged(Object^ sender, DataGridViewCellEventArgs^ e);
    void buttonExport_Click(Object^ sender, EventArgs^ e);
    void buttonRefresh_Click(Object^ sender, EventArgs^ e);
    void buttonStartAutoUpdate_Click(Object^ sender, EventArgs^ e);
    void buttonStopAutoUpdate_Click(Object^ sender, EventArgs^ e);
    void buttonToggleRegisters_Click(Object^ sender, EventArgs^ e);
    void buttonSendChanges_Click(Object^ sender, EventArgs^ e);
    void RegistersViewerForm_Load(Object^ sender, EventArgs^ e);
    void RegistersViewerForm_FormClosing(Object^ sender, FormClosingEventArgs^ e);
    void OnTimerElapsed(Object^ sender, System::Timers::ElapsedEventArgs^ e);
    void AutoUpdateData();
};