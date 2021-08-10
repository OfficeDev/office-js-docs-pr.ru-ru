---
title: Элемент DesktopSettings в файле манифеста
description: Задает исходное расположение и параметры элемента управления, которые применяются при использовании почтовой надстройки на настольном компьютере.
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: d8f35a29e14337a849f81b0becb60f761116cb48fa420119228255bb1179bb35
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089855"
---
# <a name="desktopsettings-element"></a>Элемент DesktopSettings

Задает исходное расположение и параметры элемента управления, которые применяются при использовании почтовой надстройки на настольном компьютере.

> [!IMPORTANT]
> Элемент доступен только в классических Outlook в Интернете (обычно подключенных к старым версиям локального Exchange сервера) и `DesktopSettings` Outlook 2013 Windows.

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

[Form](form.md)
