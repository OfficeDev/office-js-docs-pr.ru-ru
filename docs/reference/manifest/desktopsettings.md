---
title: Элемент DesktopSettings в файле манифеста
description: ''
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 6dfa69d407e267a1cbcfdeaad0bdf9cdf75c1465
ms.sourcegitcommit: dc42e0276007f8ab006028b9cd0cc1526c1bd100
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/14/2020
ms.locfileid: "41120644"
---
# <a name="desktopsettings-element"></a><span data-ttu-id="b1d9c-102">Элемент DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="b1d9c-102">DesktopSettings element</span></span>

<span data-ttu-id="b1d9c-103">Задает исходное расположение и параметры элемента управления, которые применяются при использовании почтовой надстройки на настольном компьютере.</span><span class="sxs-lookup"><span data-stu-id="b1d9c-103">Specifies source location and control settings that apply when your mail add-in is used on a desktop computer.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b1d9c-104">`DesktopSettings` Элемент доступен только в классическом приложении Outlook в Интернете (как правило, подключенный к старым версиям локального сервера Exchange Server) и Outlook 2013 в Windows.</span><span class="sxs-lookup"><span data-stu-id="b1d9c-104">The `DesktopSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="b1d9c-105">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="b1d9c-105">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b1d9c-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="b1d9c-106">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="b1d9c-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="b1d9c-107">Contained in</span></span>

[<span data-ttu-id="b1d9c-108">Form</span><span class="sxs-lookup"><span data-stu-id="b1d9c-108">Form</span></span>](form.md)
