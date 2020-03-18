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
# <a name="form-element"></a><span data-ttu-id="aa331-103">Элемент Form</span><span class="sxs-lookup"><span data-stu-id="aa331-103">Form element</span></span>

<span data-ttu-id="aa331-104">Параметры взаимодействия с пользователем для форм, которые почтовая надстройка будет использовать при работе на определенном устройства (настольном компьютере, планшете или телефоне).</span><span class="sxs-lookup"><span data-stu-id="aa331-104">UX settings for the forms that your mail add-in will use when running on a particular device (desktop, tablet, or phone).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="aa331-105">Элементы `DesktopSettings`, `TabletSettings`и `PhoneSettings` , Кроме того, доступны только в классическом приложении Outlook в Интернете (как правило, подключаются к предыдущим версиям локального сервера Exchange Server) и Outlook 2013 в Windows.</span><span class="sxs-lookup"><span data-stu-id="aa331-105">The `DesktopSettings`, `TabletSettings`, and `PhoneSettings` elements are available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="aa331-106">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="aa331-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="aa331-107">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="aa331-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="aa331-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="aa331-108">Contained in</span></span>

[<span data-ttu-id="aa331-109">FormSettings</span><span class="sxs-lookup"><span data-stu-id="aa331-109">FormSettings</span></span>](formsettings.md)


## <a name="can-contain"></a><span data-ttu-id="aa331-110">Может содержать</span><span class="sxs-lookup"><span data-stu-id="aa331-110">Can contain</span></span>

|<span data-ttu-id="aa331-111">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="aa331-111">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="aa331-112">DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="aa331-112">DesktopSettings</span></span>](desktopsettings.md)|
|[<span data-ttu-id="aa331-113">TabletSettings</span><span class="sxs-lookup"><span data-stu-id="aa331-113">TabletSettings</span></span>](tabletsettings.md)|
|[<span data-ttu-id="aa331-114">PhoneSettings</span><span class="sxs-lookup"><span data-stu-id="aa331-114">PhoneSettings</span></span>](phonesettings.md)|
