---
title: Элемент Form в файле манифеста
description: Параметры взаимодействия с пользователем для форм, которые почтовая надстройка будет использовать при работе на определенном устройства (настольном компьютере, планшете или телефоне).
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: 9b1696b2fecf6b07ee2a3c0a31611d4f2ad1f291
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718211"
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
