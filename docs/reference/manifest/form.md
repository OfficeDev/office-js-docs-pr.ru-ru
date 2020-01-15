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
# <a name="form-element"></a><span data-ttu-id="3b3b4-102">Элемент Form</span><span class="sxs-lookup"><span data-stu-id="3b3b4-102">Form element</span></span>

<span data-ttu-id="3b3b4-103">Параметры взаимодействия с пользователем для форм, которые почтовая надстройка будет использовать при работе на определенном устройства (настольном компьютере, планшете или телефоне).</span><span class="sxs-lookup"><span data-stu-id="3b3b4-103">UX settings for the forms that your mail add-in will use when running on a particular device (desktop, tablet, or phone).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3b3b4-104">Элементы `DesktopSettings`, `TabletSettings`и `PhoneSettings` , Кроме того, доступны только в классическом приложении Outlook в Интернете (как правило, подключаются к предыдущим версиям локального сервера Exchange Server) и Outlook 2013 в Windows.</span><span class="sxs-lookup"><span data-stu-id="3b3b4-104">The `DesktopSettings`, `TabletSettings`, and `PhoneSettings` elements are available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="3b3b4-105">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="3b3b4-105">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="3b3b4-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="3b3b4-106">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="3b3b4-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="3b3b4-107">Contained in</span></span>

[<span data-ttu-id="3b3b4-108">FormSettings</span><span class="sxs-lookup"><span data-stu-id="3b3b4-108">FormSettings</span></span>](formsettings.md)


## <a name="can-contain"></a><span data-ttu-id="3b3b4-109">Может содержать</span><span class="sxs-lookup"><span data-stu-id="3b3b4-109">Can contain</span></span>

|<span data-ttu-id="3b3b4-110">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="3b3b4-110">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="3b3b4-111">DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="3b3b4-111">DesktopSettings</span></span>](desktopsettings.md)|
|[<span data-ttu-id="3b3b4-112">TabletSettings</span><span class="sxs-lookup"><span data-stu-id="3b3b4-112">TabletSettings</span></span>](tabletsettings.md)|
|[<span data-ttu-id="3b3b4-113">PhoneSettings</span><span class="sxs-lookup"><span data-stu-id="3b3b4-113">PhoneSettings</span></span>](phonesettings.md)|
