---
title: Элемент Action в файле манифеста
description: Этот элемент указывает действие, выполняемое, когда пользователь выбирает кнопку или элемент управления меню.
ms.date: 06/08/2021
localization_priority: Normal
ms.openlocfilehash: 1ec2623ad5dbb07677735b7bcb1e39612e56984c
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936856"
---
# <a name="action-element"></a>Элемент Action

Указывает действие, выполняемое при выборе пользователем кнопки [или](control.md#button-control) [управления меню.](control.md#menu-dropdown-button-controls)

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Да  | Тип выполняемого действия|

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Описание  |
|:-----|:-----|
|  [FunctionName](#functionname) |    Указывает имя выполняемой функции. |
|  [SourceLocation](#sourcelocation) |    Указывает расположение исходного файла для этого действия. |
|  [TaskpaneId](#taskpaneid) | Определяет идентификатор для контейнера области задач. Не поддерживается Outlook надстройки.|
|  [Title](#title) | Определяет заголовок области задач. Не поддерживается Outlook надстройки.|
|  [SupportsPinning](#supportspinning) | Указывает, что область задач поддерживает закрепление (область задач остается открытой, когда пользователь выбирает другой элемент).|

## <a name="xsitype"></a>xsi:type

Этот атрибут указывает действие, которое выполняется, когда пользователь нажимает кнопку. Допустимые значения:

- `ExecuteFunction`
- `ShowTaskpane`

> [!IMPORTANT]
> Регистрация событий [почтовых ящиков](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) и [элементов](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) недоступна при **xsi:type.** `ExecuteFunction`

## <a name="functionname"></a>FunctionName

Обязательный элемент, если атрибуту **xsi:type** присвоено значение ExecuteFunction. Указывает имя выполняемой функции. Функция содержится в файле, указанном в элементе [FunctionFile](functionfile.md).

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a>SourceLocation

Необходимый **элемент, когда xsi:type** — "ShowTaskpane". Указывает расположение исходного файла для этого действия. Атрибут **resid** может быть не более 32 символов и должен быть задат к значению атрибута **id** элемента **URL** в элементе **URL-адресов** в [элементе Resources.](resources.md)

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a>TaskpaneId

Необязательный элемент, когда для атрибута **xsi:type** задано значение ShowTaskpane. Определяет идентификатор для контейнера области задач. Если у вас несколько действий ShowTaskpane и для каждого из них нужна отдельная область, используйте разные элементы **TaskpaneId**. Указывайте одинаковые элементы **TaskpaneId** для разных действий, если для последних используется одна и та же область. Когда пользователи выбирают команды, для которых используется один и тот же элемент **TaskpaneId**, контейнер области останется открытым, но оглавление области будет заменено соответствующим дочерним элементом SourceLocation элемента Action.

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

Необязательный элемент, когда для атрибута **xsi:type** задано значение ShowTaskpane. Определяет заголовок области задач для этого действия.

> [!NOTE]
> Этот элемент не поддерживается Outlook надстройки.

В следующем примере показано действие, использующее **элемент Title.** Обратите внимание, что вы не назначаете **заголовок** строке напрямую. Вместо этого вы назначите ему ИД ресурса (resid), который определяется в разделе **Ресурсы** манифеста и может быть не более 32 символов.

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="PG.Code.Url" />
    <Title resid="PG.CodeCommand.Title" />
</Action>

 ... Other markup omitted ...
<Resources>
    <bt:Images> ...
    </bt:Images>
    <bt:Urls>
        <bt:Url id="PG.Code.Url" DefaultValue="https://localhost:3000?commands=1" />
    </bt:Urls>
    <bt:ShortStrings>
        <bt:String id="PG.CodeCommand.Title" DefaultValue="Code" />
    </bt:ShortStrings>
 ... Other markup omitted ...
</Resources>
```

## <a name="supportspinning"></a>SupportsPinning

Элемент необязательный, когда для **xsi:type** задано значение ShowTaskpane. Родительские элементы [VersionOverrides](versionoverrides.md) должны иметь значение атрибута `xsi:type` `VersionOverridesV1_1`. Включите этот элемент со значением `true` для поддержки закрепления области задач. Пользователь сможет закрепить область задач, после чего она будет оставаться открытой при выборе другого элемента. Дополнительные сведения см. в статье [Реализация закрепляемой области задач в Outlook](../../outlook/pinnable-taskpane.md).

> [!IMPORTANT]
> Несмотря на то, что элемент был представлен в наборе `SupportsPinning` [требований 1.5,](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)в настоящее время он поддерживается только для Microsoft 365 абонентов с помощью следующих элементов:
>
> - Outlook 2016 или более поздней Windows (сборка 7628.1000 или более поздней)
> - Outlook 2016 или позже на Mac (сборка 16.13.503 или более поздней)
> - Современная версия Outlook в Интернете

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
