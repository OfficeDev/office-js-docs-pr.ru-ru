---
title: Элемент TabletSettingst в файле манифеста
description: ''
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: bf11dcfec4dfe2c40764722d23c7a69c289bba65
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163866"
---
# <a name="tabletsettings-element"></a><span data-ttu-id="99229-102">Элемент TabletSettingst</span><span class="sxs-lookup"><span data-stu-id="99229-102">TabletSettings element</span></span>

<span data-ttu-id="99229-103">Задает параметры управления, которые применяются при использовании вашей почтовой надстройки на планшете.</span><span class="sxs-lookup"><span data-stu-id="99229-103">Specifies control settings that apply when your mail add-in is used on a tablet.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="99229-104">`TabletSettings` Элемент доступен только в классическом приложении Outlook в Интернете (как правило, подключенный к старым версиям локального сервера Exchange Server) и Outlook 2013 в Windows.</span><span class="sxs-lookup"><span data-stu-id="99229-104">The `TabletSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="99229-105">Для поддержки Outlook в Android и iOS, ознакомьтесь со статьей надстройки [для Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span><span class="sxs-lookup"><span data-stu-id="99229-105">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span></span>

<span data-ttu-id="99229-106">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="99229-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="99229-107">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="99229-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="99229-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="99229-108">Contained in</span></span>

[<span data-ttu-id="99229-109">Form</span><span class="sxs-lookup"><span data-stu-id="99229-109">Form</span></span>](form.md)

