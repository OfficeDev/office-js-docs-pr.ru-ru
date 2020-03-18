---
title: Элемент DesktopSettings в файле манифеста
description: Задает исходное расположение и параметры элемента управления, которые применяются при использовании почтовой надстройки на настольном компьютере.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: d48532482fc71fec2a96133ee8e813cae798613f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718358"
---
# <a name="desktopsettings-element"></a><span data-ttu-id="2f4ab-103">Элемент DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="2f4ab-103">DesktopSettings element</span></span>

<span data-ttu-id="2f4ab-104">Задает исходное расположение и параметры элемента управления, которые применяются при использовании почтовой надстройки на настольном компьютере.</span><span class="sxs-lookup"><span data-stu-id="2f4ab-104">Specifies source location and control settings that apply when your mail add-in is used on a desktop computer.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2f4ab-105">`DesktopSettings` Элемент доступен только в классическом приложении Outlook в Интернете (как правило, подключенный к старым версиям локального сервера Exchange Server) и Outlook 2013 в Windows.</span><span class="sxs-lookup"><span data-stu-id="2f4ab-105">The `DesktopSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="2f4ab-106">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="2f4ab-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="2f4ab-107">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="2f4ab-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="2f4ab-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="2f4ab-108">Contained in</span></span>

[<span data-ttu-id="2f4ab-109">Form</span><span class="sxs-lookup"><span data-stu-id="2f4ab-109">Form</span></span>](form.md)
