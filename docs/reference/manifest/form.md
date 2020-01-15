---
title: Элемент Form в файле манифеста
description: ''
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: d545d471e007f0077a8310b0b847bbbf99a8f7ac
ms.sourcegitcommit: dc42e0276007f8ab006028b9cd0cc1526c1bd100
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/14/2020
ms.locfileid: "41120651"
---
# <a name="form-element"></a>Элемент Form

Параметры взаимодействия с пользователем для форм, которые почтовая надстройка будет использовать при работе на определенном устройства (настольном компьютере, планшете или телефоне).

> [!IMPORTANT]
> Элементы `DesktopSettings`, `TabletSettings`и `PhoneSettings` , Кроме того, доступны только в классическом приложении Outlook в Интернете (как правило, подключаются к предыдущим версиям локального сервера Exchange Server) и Outlook 2013 в Windows.

**Тип надстройки:** почтовая

## <a name="syntax"></a>Синтаксис

```XML
<Form xsi:type="ItemRead">
   <!--website.html is a placeholder for your own add-in website.-->
   <DesktopSettings>
      <SourceLocation DefaultValue="https://website.html" />
      <!--RequestedHeight must be between 240px to 800px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
   </DesktopSettings>
   <TabletSettings>
      <SourceLocation DefaultValue="https://website.html" />
      <!--RequestedHeight must be between 240px to 800px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
   </TabletSettings>
   <PhoneSettings>
      <SourceLocation DefaultValue="https://website.html" />
   </PhoneSettings>
</Form>
```

## <a name="contained-in"></a>Содержится в

[FormSettings](formsettings.md)


## <a name="can-contain"></a>Может содержать

|**Элемент**|
|:-----|
|[DesktopSettings](desktopsettings.md)|
|[TabletSettings](tabletsettings.md)|
|[PhoneSettings](phonesettings.md)|
