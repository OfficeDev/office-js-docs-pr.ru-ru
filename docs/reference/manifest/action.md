---
title: Элемент Action в файле манифеста
description: ''
ms.date: 11/14/2018
ms.openlocfilehash: 04c081a02768446fcf587b8b6a7c4e1dcd66012f
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433273"
---
# <a name="action-element"></a>Элемент Action

Указывает действие, которое необходимо выполнить, когда пользователь выбирает элемент управления [Button](control.md#button-control) или [Menu](control.md#menu-dropdown-button-controls).

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Да  | Тип выполняемого действия|

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Описание  |
|:-----|:-----|
|  [FunctionName](#functionname) |    Указывает имя выполняемой функции. |
|  [SourceLocation](#sourcelocation) |    Указывает расположение исходного файла для этого действия. |
|  [TaskpaneId](#taskpaneid) | Определяет идентификатор контейнера области задач.|
|  [Title](#title) | Определяет заголовок области задач.|
|  [SupportsPinning](#supportspinning) | Указывает, что область задач поддерживает закрепление (область задач остается открытой, когда пользователь выбирает другой элемент).|
  

## <a name="xsitype"></a>xsi:type

Этот атрибут указывает действие, которое выполняется, когда пользователь нажимает кнопку. Допустимые значения:

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a>FunctionName

Обязательный элемент, если атрибуту **xsi:type** присвоено значение ExecuteFunction. Указывает имя выполняемой функции. Функция содержится в файле, указанном в элементе [FunctionFile](functionfile.md).

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a>SourceLocation

Обязательный элемент, если атрибуту **xsi:type** присвоено значение ShowTaskpane. Указывает расположение исходного файла для этого действия. Атрибуту **resid** должно быть присвоено значение атрибута **id** элемента **Url** в элементе **Urls**, включенном в элемент [Resources](resources.md).

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a>TaskpaneId

Элемент необязательный, когда для  **xsi:type** задано значение ShowTaskpane. Определяет идентификатор контейнера области задач. Если у вас несколько действий ShowTaskpane и для каждого из них нужна отдельная область, используйте разные элементы **TaskpaneId**. Указывайте одинаковые элементы **TaskpaneId** для разных действий, если для последних используется одна и та же область. Когда пользователи выбирают команды, для которых используется один и тот же элемент **TaskpaneId**, контейнер области останется открытым, но содержимое области будет заменено соответствующим дочерним элементом SourceLocation элемента Action.

> [!NOTE]
> Этот элемент не поддерживается в Outlook.

В следующем примере показаны два действия, для которых используется один и тот же элемент **TaskpaneId**.

```xml
<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="aTaskPaneUrl" />
</Action>

<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="anotherTaskPaneUrl" />
</Action>
```  

В следующих примерах показаны два действия, использующие другой элемент **TaskpaneId**. Чтобы увидеть эти примеры в контексте, ознакомьтесь с [примером команд простых надстроек](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).

```xml
<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID1</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane1.Url" />
</Action>

<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID2</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane2.Url" />
</Action>
```  

```xml
<bt:Urls>
   <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
   <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
</bt:Urls>
```  

## <a name="title"></a>Должность

Элемент необязательный, когда для  **xsi:type** задано значение ShowTaskpane. Определяет заголовок области задач для этого действия.

В приведенных ниже примерах показаны два действия, для которых используется элемент **Title**.

```xml
<Action xsi:type="ShowTaskpane">
<TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
<SourceLocation resid="PG.Code.Url" />
<Title resid="PG.CodeCommand.Title" />
</Action>
```

```xml
<Action xsi:type="ShowTaskpane">
<SourceLocation resid="PG.Run.Url" />
<Title resid="PG.RunCommand.Title" />
</Action>
```

```xml
<bt:Urls>
<bt:Url id="PG.Code.Url" DefaultValue="https://localhost:3000?commands=1" />
<bt:Url id="PG.Run.Url" DefaultValue="https://localhost:3000/run.html" />
</bt:Urls>
```

```xml
<bt:ShortStrings>
<bt:String id="PG.CodeCommand.Title" DefaultValue="Code" />
<bt:String id="PG.RunCommand.Title" DefaultValue="Run" />
</bt:ShortStrings>
```

## <a name="supportspinning"></a>SupportsPinning

Элемент необязательный, когда для **xsi:type** задано значение ShowTaskpane. Родительские элементы [VersionOverrides](versionoverrides.md) должны иметь значение атрибута `xsi:type` `VersionOverridesV1_1`. Включите этот элемент со значением `true` для поддержки закрепления области задач. Пользователь сможет закрепить область задач, после чего она будет оставаться открытой при выборе другого элемента. Дополнительные сведения см. в статье [Реализация закрепляемой области задач в Outlook](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).

> [!NOTE]
> В настоящее время элемент SupportsPinning поддерживается только в Outlook 2016 для Windows (сборка 7628.1000 или более поздней версии).

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
