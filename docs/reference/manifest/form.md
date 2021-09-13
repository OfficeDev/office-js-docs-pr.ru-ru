---
title: Элемент Form в файле манифеста
description: Параметры взаимодействия с пользователем для форм, которые почтовая надстройка будет использовать при работе на определенном устройства (настольном компьютере, планшете или телефоне).
ms.date: 04/09/2020
ms.localizationpriority: medium
ms.openlocfilehash: ecea978cd87a9feed08ba621c9377525acaf4e3a
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151133"
---
# <a name="form-element"></a>Элемент Form

Параметры взаимодействия с пользователем для форм, которые почтовая надстройка будет использовать при работе на определенном устройства (настольном компьютере, планшете или телефоне).

> [!IMPORTANT]
> Элементы и элементы доступны только в классических Outlook в Интернете (обычно подключенных к старым версиям локального Exchange сервера) и `DesktopSettings` `TabletSettings` Outlook `PhoneSettings` 2013 на Windows.

**Тип надстройки:** почтовая

## <a name="syntax"></a>Синтаксис

```XML
<Form xsi:type="ItemRead">
   <!--https://MyDomain.com/website.html is a placeholder for your own add-in website.-->
   <DesktopSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </DesktopSettings>
   <TabletSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </TabletSettings>
   <PhoneSettings>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
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
