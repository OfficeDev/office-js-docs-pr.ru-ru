# <a name="host-element"></a>Элемент Host

Определяет тип приложения Office, в котором следует активировать надстройку.

> [!IMPORTANT] 
> Синтаксис элемента **Host** зависит от того, задается ли элемент в [базовом манифесте](#basic-manifest) или в узле [VersionOverrides](#versionoverrides-node). Однако функциональность в обоих случаях одинакова.  

## <a name="basic-manifest"></a>Базовый манифест

Если основное приложение задается в базовом манифесте (в разделе [OfficeApp](officeapp.md)), то его тип определяется атрибутом `Name`.   

### <a name="attributes"></a>Атрибуты

| Атрибут     | Тип   | Обязательный | Описание                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [Имя](#name) | строка | обязательный | Имя типа основного приложения Office. |

### <a name="name"></a>Имя
Определяет тип основного приложения, для которого предназначена эта надстройка. Поддерживаются такие значения:

- `Document` (Word)
- `Database` (Access)
- `Mailbox` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Project` (Project)
- `Workbook` (Excel)

### <a name="example"></a>Пример
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a>Узел VersionOverrides
Если основной элемент задается в узле [VersionOverrides](versionoverrides.md), его тип определяет атрибут `xsi:type`. 

### <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Да  | Описывает приложение Office, к которому применяются эти параметры.|

### <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [DesktopFormFactor](desktopformfactor.md)    |  Да   |  Определяет параметры классического форм-фактора. |
|  [MobileFormFactor](mobileformfactor.md)    |  Нет   |  Определяет параметры форм-фактора мобильного устройства. **Примечание.** Этот элемент поддерживается только в Outlook для iOS. |
|  [AllFormFactors](allformfactors.md)    |  Нет   |  Определяет параметры всех форм-факторов. Используется только пользовательскими функциями в Excel. |

### <a name="xsitype"></a>xsi:type

Указывает, к какому основному приложению Office (Word, Excel, PowerPoint, Outlook, OneNote) применяются содержащиеся параметры. Допустимые значения:

- `Document` (Word)
- `MailHost` (Outlook)    
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Workbook` (Excel)

## <a name="host-example"></a>Пример основного приложения 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
