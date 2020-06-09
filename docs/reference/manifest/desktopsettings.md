---
title: Элемент DesktopSettings в файле манифеста
description: Задает исходное расположение и параметры элемента управления, которые применяются при использовании почтовой надстройки на настольном компьютере.
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 50201080d8be3c8943d16730c34a4bac236d7b90
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612278"
---
# <a name="desktopsettings-element"></a><span data-ttu-id="bd9fe-103">Элемент DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="bd9fe-103">DesktopSettings element</span></span>

<span data-ttu-id="bd9fe-104">Задает исходное расположение и параметры элемента управления, которые применяются при использовании почтовой надстройки на настольном компьютере.</span><span class="sxs-lookup"><span data-stu-id="bd9fe-104">Specifies source location and control settings that apply when your mail add-in is used on a desktop computer.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bd9fe-105">`DesktopSettings`Элемент доступен только в классическом приложении Outlook в Интернете (как правило, подключенный к старым версиям локального сервера Exchange Server) и Outlook 2013 в Windows.</span><span class="sxs-lookup"><span data-stu-id="bd9fe-105">The `DesktopSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="bd9fe-106">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="bd9fe-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="bd9fe-107">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="bd9fe-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="bd9fe-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="bd9fe-108">Contained in</span></span>

[<span data-ttu-id="bd9fe-109">Form</span><span class="sxs-lookup"><span data-stu-id="bd9fe-109">Form</span></span>](form.md)
