---
title: Office. Context. Mailbox — Предварительная версия набора обязательных элементов
description: ''
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: 864c4f2931762ff6d8a02abb8da1a03e1abcab80
ms.sourcegitcommit: 44f1a4a3e1ae3c33d7d5fabcee14b84af94e03da
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/03/2019
ms.locfileid: "39670120"
---
# <a name="mailbox"></a><span data-ttu-id="778f7-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="778f7-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="778f7-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="778f7-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="778f7-104">Предоставляет для Microsoft Outlook доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="778f7-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="778f7-105">Требования</span><span class="sxs-lookup"><span data-stu-id="778f7-105">Requirements</span></span>

|<span data-ttu-id="778f7-106">Требование</span><span class="sxs-lookup"><span data-stu-id="778f7-106">Requirement</span></span>| <span data-ttu-id="778f7-107">Значение</span><span class="sxs-lookup"><span data-stu-id="778f7-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="778f7-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="778f7-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778f7-109">1.0</span><span class="sxs-lookup"><span data-stu-id="778f7-109">1.0</span></span>|
|[<span data-ttu-id="778f7-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="778f7-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778f7-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="778f7-111">Restricted</span></span>|
|[<span data-ttu-id="778f7-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="778f7-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778f7-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="778f7-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="778f7-114">Properties</span></span>

| <span data-ttu-id="778f7-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="778f7-115">Property</span></span> | <span data-ttu-id="778f7-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="778f7-116">Minimum</span></span><br><span data-ttu-id="778f7-117">уровень разрешения</span><span class="sxs-lookup"><span data-stu-id="778f7-117">permission level</span></span> | <span data-ttu-id="778f7-118">Способов</span><span class="sxs-lookup"><span data-stu-id="778f7-118">Modes</span></span> | <span data-ttu-id="778f7-119">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="778f7-119">Return type</span></span> | <span data-ttu-id="778f7-120">Минимальные</span><span class="sxs-lookup"><span data-stu-id="778f7-120">Minimum</span></span><br><span data-ttu-id="778f7-121">набор требований</span><span class="sxs-lookup"><span data-stu-id="778f7-121">requirement set</span></span> |
|---|---|---|---|---|
| [<span data-ttu-id="778f7-122">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="778f7-122">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="778f7-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-123">ReadItem</span></span> | <span data-ttu-id="778f7-124">Создание</span><span class="sxs-lookup"><span data-stu-id="778f7-124">Compose</span></span><br><span data-ttu-id="778f7-125">Чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-125">Read</span></span> | <span data-ttu-id="778f7-126">Строка</span><span class="sxs-lookup"><span data-stu-id="778f7-126">String</span></span> | <span data-ttu-id="778f7-127">1.0</span><span class="sxs-lookup"><span data-stu-id="778f7-127">1.0</span></span> |
| [<span data-ttu-id="778f7-128">мастеркатегориес</span><span class="sxs-lookup"><span data-stu-id="778f7-128">masterCategories</span></span>](#mastercategories-mastercategories) | <span data-ttu-id="778f7-129">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="778f7-129">ReadWriteMailbox</span></span> | <span data-ttu-id="778f7-130">Создание</span><span class="sxs-lookup"><span data-stu-id="778f7-130">Compose</span></span><br><span data-ttu-id="778f7-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-131">Read</span></span> | [<span data-ttu-id="778f7-132">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="778f7-132">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories) | <span data-ttu-id="778f7-133">1.8</span><span class="sxs-lookup"><span data-stu-id="778f7-133">1.8</span></span> |
| [<span data-ttu-id="778f7-134">restUrl</span><span class="sxs-lookup"><span data-stu-id="778f7-134">restUrl</span></span>](#resturl-string) | <span data-ttu-id="778f7-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-135">ReadItem</span></span> | <span data-ttu-id="778f7-136">Создание</span><span class="sxs-lookup"><span data-stu-id="778f7-136">Compose</span></span><br><span data-ttu-id="778f7-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-137">Read</span></span> | <span data-ttu-id="778f7-138">Строка</span><span class="sxs-lookup"><span data-stu-id="778f7-138">String</span></span> | <span data-ttu-id="778f7-139">1.5</span><span class="sxs-lookup"><span data-stu-id="778f7-139">1.5</span></span> |

##### <a name="methods"></a><span data-ttu-id="778f7-140">Методы</span><span class="sxs-lookup"><span data-stu-id="778f7-140">Methods</span></span>

| <span data-ttu-id="778f7-141">Метод</span><span class="sxs-lookup"><span data-stu-id="778f7-141">Method</span></span> | <span data-ttu-id="778f7-142">Минимальные</span><span class="sxs-lookup"><span data-stu-id="778f7-142">Minimum</span></span><br><span data-ttu-id="778f7-143">уровень разрешения</span><span class="sxs-lookup"><span data-stu-id="778f7-143">permission level</span></span> | <span data-ttu-id="778f7-144">Способов</span><span class="sxs-lookup"><span data-stu-id="778f7-144">Modes</span></span> | <span data-ttu-id="778f7-145">Минимальные</span><span class="sxs-lookup"><span data-stu-id="778f7-145">Minimum</span></span><br><span data-ttu-id="778f7-146">набор требований</span><span class="sxs-lookup"><span data-stu-id="778f7-146">requirement set</span></span> |
|---|---|---|---|
| [<span data-ttu-id="778f7-147">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="778f7-147">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="778f7-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-148">ReadItem</span></span> | <span data-ttu-id="778f7-149">Создание</span><span class="sxs-lookup"><span data-stu-id="778f7-149">Compose</span></span><br><span data-ttu-id="778f7-150">Чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-150">Read</span></span> | <span data-ttu-id="778f7-151">1.5</span><span class="sxs-lookup"><span data-stu-id="778f7-151">1.5</span></span> |
| [<span data-ttu-id="778f7-152">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="778f7-152">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="778f7-153">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="778f7-153">Restricted</span></span> | <span data-ttu-id="778f7-154">Создание</span><span class="sxs-lookup"><span data-stu-id="778f7-154">Compose</span></span><br><span data-ttu-id="778f7-155">Чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-155">Read</span></span> | <span data-ttu-id="778f7-156">1.3</span><span class="sxs-lookup"><span data-stu-id="778f7-156">1.3</span></span> |
| [<span data-ttu-id="778f7-157">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="778f7-157">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="778f7-158">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-158">ReadItem</span></span> | <span data-ttu-id="778f7-159">Создание</span><span class="sxs-lookup"><span data-stu-id="778f7-159">Compose</span></span><br><span data-ttu-id="778f7-160">Чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-160">Read</span></span> | <span data-ttu-id="778f7-161">1.0</span><span class="sxs-lookup"><span data-stu-id="778f7-161">1.0</span></span> |
| [<span data-ttu-id="778f7-162">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="778f7-162">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="778f7-163">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="778f7-163">Restricted</span></span> | <span data-ttu-id="778f7-164">Создание</span><span class="sxs-lookup"><span data-stu-id="778f7-164">Compose</span></span><br><span data-ttu-id="778f7-165">Чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-165">Read</span></span> | <span data-ttu-id="778f7-166">1.3</span><span class="sxs-lookup"><span data-stu-id="778f7-166">1.3</span></span> |
| [<span data-ttu-id="778f7-167">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="778f7-167">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="778f7-168">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-168">ReadItem</span></span> | <span data-ttu-id="778f7-169">Создание</span><span class="sxs-lookup"><span data-stu-id="778f7-169">Compose</span></span><br><span data-ttu-id="778f7-170">Чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-170">Read</span></span> | <span data-ttu-id="778f7-171">1.0</span><span class="sxs-lookup"><span data-stu-id="778f7-171">1.0</span></span> |
| [<span data-ttu-id="778f7-172">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="778f7-172">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="778f7-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-173">ReadItem</span></span> | <span data-ttu-id="778f7-174">Создание</span><span class="sxs-lookup"><span data-stu-id="778f7-174">Compose</span></span><br><span data-ttu-id="778f7-175">Чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-175">Read</span></span> | <span data-ttu-id="778f7-176">1.0</span><span class="sxs-lookup"><span data-stu-id="778f7-176">1.0</span></span> |
| [<span data-ttu-id="778f7-177">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="778f7-177">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="778f7-178">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-178">ReadItem</span></span> | <span data-ttu-id="778f7-179">Создание</span><span class="sxs-lookup"><span data-stu-id="778f7-179">Compose</span></span><br><span data-ttu-id="778f7-180">Чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-180">Read</span></span> | <span data-ttu-id="778f7-181">1.0</span><span class="sxs-lookup"><span data-stu-id="778f7-181">1.0</span></span> |
| [<span data-ttu-id="778f7-182">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="778f7-182">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="778f7-183">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-183">ReadItem</span></span> | <span data-ttu-id="778f7-184">Чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-184">Read</span></span> | <span data-ttu-id="778f7-185">1.0</span><span class="sxs-lookup"><span data-stu-id="778f7-185">1.0</span></span> |
| [<span data-ttu-id="778f7-186">дисплайневмессажеформ</span><span class="sxs-lookup"><span data-stu-id="778f7-186">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="778f7-187">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-187">ReadItem</span></span> | <span data-ttu-id="778f7-188">Создание</span><span class="sxs-lookup"><span data-stu-id="778f7-188">Compose</span></span><br><span data-ttu-id="778f7-189">Чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-189">Read</span></span> | <span data-ttu-id="778f7-190">1.6</span><span class="sxs-lookup"><span data-stu-id="778f7-190">1.6</span></span> |
| [<span data-ttu-id="778f7-191">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="778f7-191">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="778f7-192">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-192">ReadItem</span></span> | <span data-ttu-id="778f7-193">Создание</span><span class="sxs-lookup"><span data-stu-id="778f7-193">Compose</span></span><br><span data-ttu-id="778f7-194">Чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-194">Read</span></span> | <span data-ttu-id="778f7-195">1.5</span><span class="sxs-lookup"><span data-stu-id="778f7-195">1.5</span></span> |
| [<span data-ttu-id="778f7-196">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="778f7-196">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="778f7-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-197">ReadItem</span></span> | <span data-ttu-id="778f7-198">Создание</span><span class="sxs-lookup"><span data-stu-id="778f7-198">Compose</span></span><br><span data-ttu-id="778f7-199">Чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-199">Read</span></span> | <span data-ttu-id="778f7-200">1.3</span><span class="sxs-lookup"><span data-stu-id="778f7-200">1.3</span></span><br><span data-ttu-id="778f7-201">1.0</span><span class="sxs-lookup"><span data-stu-id="778f7-201">1.0</span></span> |
| [<span data-ttu-id="778f7-202">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="778f7-202">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="778f7-203">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-203">ReadItem</span></span> | <span data-ttu-id="778f7-204">Создание</span><span class="sxs-lookup"><span data-stu-id="778f7-204">Compose</span></span><br><span data-ttu-id="778f7-205">Чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-205">Read</span></span> | <span data-ttu-id="778f7-206">1.0</span><span class="sxs-lookup"><span data-stu-id="778f7-206">1.0</span></span> |
| [<span data-ttu-id="778f7-207">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="778f7-207">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="778f7-208">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="778f7-208">ReadWriteMailbox</span></span> | <span data-ttu-id="778f7-209">Создание</span><span class="sxs-lookup"><span data-stu-id="778f7-209">Compose</span></span><br><span data-ttu-id="778f7-210">Чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-210">Read</span></span> | <span data-ttu-id="778f7-211">1.0</span><span class="sxs-lookup"><span data-stu-id="778f7-211">1.0</span></span> |
| [<span data-ttu-id="778f7-212">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="778f7-212">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="778f7-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-213">ReadItem</span></span> | <span data-ttu-id="778f7-214">Создание</span><span class="sxs-lookup"><span data-stu-id="778f7-214">Compose</span></span><br><span data-ttu-id="778f7-215">Чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-215">Read</span></span> | <span data-ttu-id="778f7-216">1.5</span><span class="sxs-lookup"><span data-stu-id="778f7-216">1.5</span></span> |

##### <a name="events"></a><span data-ttu-id="778f7-217">События</span><span class="sxs-lookup"><span data-stu-id="778f7-217">Events</span></span>

<span data-ttu-id="778f7-218">Вы можете подписаться на следующие события и отписаться на них, используя [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) и [removeHandlerAsync](#removehandlerasynceventtype-options-callback) соответственно.</span><span class="sxs-lookup"><span data-stu-id="778f7-218">You can subscribe to and unsubscribe from the following events using [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) and [removeHandlerAsync](#removehandlerasynceventtype-options-callback) respectively.</span></span>

| <span data-ttu-id="778f7-219">Событие</span><span class="sxs-lookup"><span data-stu-id="778f7-219">Event</span></span> | <span data-ttu-id="778f7-220">Описание</span><span class="sxs-lookup"><span data-stu-id="778f7-220">Description</span></span> | <span data-ttu-id="778f7-221">Минимальные</span><span class="sxs-lookup"><span data-stu-id="778f7-221">Minimum</span></span><br><span data-ttu-id="778f7-222">набор требований</span><span class="sxs-lookup"><span data-stu-id="778f7-222">requirement set</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="778f7-223">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="778f7-223">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="778f7-224">1.5</span><span class="sxs-lookup"><span data-stu-id="778f7-224">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="778f7-225">Тема Office в почтовом ящике изменилась.</span><span class="sxs-lookup"><span data-stu-id="778f7-225">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="778f7-226">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="778f7-226">Preview</span></span> |

### <a name="namespaces"></a><span data-ttu-id="778f7-227">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="778f7-227">Namespaces</span></span>

<span data-ttu-id="778f7-228">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="778f7-228">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="778f7-229">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="778f7-229">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="778f7-230">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="778f7-230">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

## <a name="property-details"></a><span data-ttu-id="778f7-231">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="778f7-231">Property details</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="778f7-232">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="778f7-232">ewsUrl: String</span></span>

<span data-ttu-id="778f7-p101">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="778f7-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="778f7-235">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="778f7-235">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="778f7-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="778f7-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="778f7-238">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="778f7-238">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="778f7-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="778f7-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="778f7-241">Тип</span><span class="sxs-lookup"><span data-stu-id="778f7-241">Type</span></span>

*   <span data-ttu-id="778f7-242">String</span><span class="sxs-lookup"><span data-stu-id="778f7-242">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="778f7-243">Требования</span><span class="sxs-lookup"><span data-stu-id="778f7-243">Requirements</span></span>

|<span data-ttu-id="778f7-244">Требование</span><span class="sxs-lookup"><span data-stu-id="778f7-244">Requirement</span></span>| <span data-ttu-id="778f7-245">Значение</span><span class="sxs-lookup"><span data-stu-id="778f7-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="778f7-246">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="778f7-246">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778f7-247">1.0</span><span class="sxs-lookup"><span data-stu-id="778f7-247">1.0</span></span>|
|[<span data-ttu-id="778f7-248">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="778f7-248">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778f7-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-249">ReadItem</span></span>|
|[<span data-ttu-id="778f7-250">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="778f7-250">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778f7-251">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-251">Compose or Read</span></span>|

<br>

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategories"></a><span data-ttu-id="778f7-252">Мастеркатегориес: [мастеркатегориес](/javascript/api/outlook/office.mastercategories)</span><span class="sxs-lookup"><span data-stu-id="778f7-252">masterCategories: [MasterCategories](/javascript/api/outlook/office.mastercategories)</span></span>

<span data-ttu-id="778f7-253">Получает объект, предоставляющий методы для управления главным списком категорий в этом почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="778f7-253">Gets an object that provides methods to manage the categories master list on this mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="778f7-254">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="778f7-254">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="778f7-255">Тип</span><span class="sxs-lookup"><span data-stu-id="778f7-255">Type</span></span>

*   [<span data-ttu-id="778f7-256">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="778f7-256">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

##### <a name="requirements"></a><span data-ttu-id="778f7-257">Требования</span><span class="sxs-lookup"><span data-stu-id="778f7-257">Requirements</span></span>

|<span data-ttu-id="778f7-258">Требование</span><span class="sxs-lookup"><span data-stu-id="778f7-258">Requirement</span></span>| <span data-ttu-id="778f7-259">Значение</span><span class="sxs-lookup"><span data-stu-id="778f7-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="778f7-260">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="778f7-260">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778f7-261">1.8</span><span class="sxs-lookup"><span data-stu-id="778f7-261">1.8</span></span> |
|[<span data-ttu-id="778f7-262">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="778f7-262">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778f7-263">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="778f7-263">ReadWriteMailbox</span></span> |
|[<span data-ttu-id="778f7-264">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="778f7-264">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778f7-265">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-265">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="778f7-266">Пример</span><span class="sxs-lookup"><span data-stu-id="778f7-266">Example</span></span>

<span data-ttu-id="778f7-267">В этом примере показано получение сводного списка категорий для этого почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="778f7-267">This example gets the categories master list for this mailbox.</span></span>

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Master categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="778f7-268">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="778f7-268">restUrl: String</span></span>

<span data-ttu-id="778f7-269">Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="778f7-269">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="778f7-270">С помощью значения `restUrl` можно выполнять вызовы [REST API](/outlook/rest/) для почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="778f7-270">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="778f7-271">Тип</span><span class="sxs-lookup"><span data-stu-id="778f7-271">Type</span></span>

*   <span data-ttu-id="778f7-272">String</span><span class="sxs-lookup"><span data-stu-id="778f7-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="778f7-273">Требования</span><span class="sxs-lookup"><span data-stu-id="778f7-273">Requirements</span></span>

|<span data-ttu-id="778f7-274">Требование</span><span class="sxs-lookup"><span data-stu-id="778f7-274">Requirement</span></span>| <span data-ttu-id="778f7-275">Значение</span><span class="sxs-lookup"><span data-stu-id="778f7-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="778f7-276">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="778f7-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778f7-277">1.5</span><span class="sxs-lookup"><span data-stu-id="778f7-277">1.5</span></span> |
|[<span data-ttu-id="778f7-278">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="778f7-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778f7-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-279">ReadItem</span></span>|
|[<span data-ttu-id="778f7-280">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="778f7-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778f7-281">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-281">Compose or Read</span></span>|

## <a name="method-details"></a><span data-ttu-id="778f7-282">Сведения о методе</span><span class="sxs-lookup"><span data-stu-id="778f7-282">Method details</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="778f7-283">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="778f7-283">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="778f7-284">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="778f7-284">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="778f7-285">В настоящее время поддерживаются типы событий `Office.EventType.ItemChanged` и `Office.EventType.OfficeThemeChanged`.</span><span class="sxs-lookup"><span data-stu-id="778f7-285">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="778f7-286">Параметры</span><span class="sxs-lookup"><span data-stu-id="778f7-286">Parameters</span></span>

| <span data-ttu-id="778f7-287">Имя</span><span class="sxs-lookup"><span data-stu-id="778f7-287">Name</span></span> | <span data-ttu-id="778f7-288">Тип</span><span class="sxs-lookup"><span data-stu-id="778f7-288">Type</span></span> | <span data-ttu-id="778f7-289">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="778f7-289">Attributes</span></span> | <span data-ttu-id="778f7-290">Описание</span><span class="sxs-lookup"><span data-stu-id="778f7-290">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="778f7-291">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="778f7-291">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="778f7-292">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="778f7-292">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="778f7-293">Function</span><span class="sxs-lookup"><span data-stu-id="778f7-293">Function</span></span> || <span data-ttu-id="778f7-p104">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="778f7-p104">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="778f7-297">Объект</span><span class="sxs-lookup"><span data-stu-id="778f7-297">Object</span></span> | <span data-ttu-id="778f7-298">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="778f7-298">&lt;optional&gt;</span></span> | <span data-ttu-id="778f7-299">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="778f7-299">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="778f7-300">Object</span><span class="sxs-lookup"><span data-stu-id="778f7-300">Object</span></span> | <span data-ttu-id="778f7-301">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="778f7-301">&lt;optional&gt;</span></span> | <span data-ttu-id="778f7-302">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="778f7-302">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="778f7-303">функция</span><span class="sxs-lookup"><span data-stu-id="778f7-303">function</span></span>| <span data-ttu-id="778f7-304">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="778f7-304">&lt;optional&gt;</span></span>|<span data-ttu-id="778f7-305">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="778f7-305">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778f7-306">Требования</span><span class="sxs-lookup"><span data-stu-id="778f7-306">Requirements</span></span>

|<span data-ttu-id="778f7-307">Требование</span><span class="sxs-lookup"><span data-stu-id="778f7-307">Requirement</span></span>| <span data-ttu-id="778f7-308">Значение</span><span class="sxs-lookup"><span data-stu-id="778f7-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="778f7-309">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="778f7-309">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778f7-310">1.5</span><span class="sxs-lookup"><span data-stu-id="778f7-310">1.5</span></span> |
|[<span data-ttu-id="778f7-311">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="778f7-311">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778f7-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-312">ReadItem</span></span> |
|[<span data-ttu-id="778f7-313">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="778f7-313">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778f7-314">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-314">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="778f7-315">Пример</span><span class="sxs-lookup"><span data-stu-id="778f7-315">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error.
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item.
  loadProps(Office.context.mailbox.item);
}
```

<br>

---
---

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="778f7-316">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="778f7-316">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="778f7-317">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="778f7-317">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="778f7-318">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="778f7-318">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="778f7-p105">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="778f7-p105">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="778f7-321">Параметры</span><span class="sxs-lookup"><span data-stu-id="778f7-321">Parameters</span></span>

|<span data-ttu-id="778f7-322">Имя</span><span class="sxs-lookup"><span data-stu-id="778f7-322">Name</span></span>| <span data-ttu-id="778f7-323">Тип</span><span class="sxs-lookup"><span data-stu-id="778f7-323">Type</span></span>| <span data-ttu-id="778f7-324">Описание</span><span class="sxs-lookup"><span data-stu-id="778f7-324">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="778f7-325">Строка</span><span class="sxs-lookup"><span data-stu-id="778f7-325">String</span></span>|<span data-ttu-id="778f7-326">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="778f7-326">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="778f7-327">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="778f7-327">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="778f7-328">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="778f7-328">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778f7-329">Требования</span><span class="sxs-lookup"><span data-stu-id="778f7-329">Requirements</span></span>

|<span data-ttu-id="778f7-330">Требование</span><span class="sxs-lookup"><span data-stu-id="778f7-330">Requirement</span></span>| <span data-ttu-id="778f7-331">Значение</span><span class="sxs-lookup"><span data-stu-id="778f7-331">Value</span></span>|
|---|---|
|[<span data-ttu-id="778f7-332">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="778f7-332">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778f7-333">1.3</span><span class="sxs-lookup"><span data-stu-id="778f7-333">1.3</span></span>|
|[<span data-ttu-id="778f7-334">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="778f7-334">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778f7-335">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="778f7-335">Restricted</span></span>|
|[<span data-ttu-id="778f7-336">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="778f7-336">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778f7-337">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-337">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="778f7-338">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="778f7-338">Returns:</span></span>

<span data-ttu-id="778f7-339">Тип: String</span><span class="sxs-lookup"><span data-stu-id="778f7-339">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="778f7-340">Пример</span><span class="sxs-lookup"><span data-stu-id="778f7-340">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="778f7-341">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="778f7-341">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="778f7-342">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="778f7-342">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="778f7-p106">Почтовое приложение для классической версии Outlook или версии в Интернете может использовать разные часовые пояса для дат и времени. Классическое приложение Outlook использует часовой пояс клиентского компьютера. Outlook в Интернете использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="778f7-p106">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="778f7-p107">Если почтовое приложение работает в классическом клиенте Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook в Интернете, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="778f7-p107">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="778f7-348">Параметры</span><span class="sxs-lookup"><span data-stu-id="778f7-348">Parameters</span></span>

|<span data-ttu-id="778f7-349">Имя</span><span class="sxs-lookup"><span data-stu-id="778f7-349">Name</span></span>| <span data-ttu-id="778f7-350">Тип</span><span class="sxs-lookup"><span data-stu-id="778f7-350">Type</span></span>| <span data-ttu-id="778f7-351">Описание</span><span class="sxs-lookup"><span data-stu-id="778f7-351">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="778f7-352">Date</span><span class="sxs-lookup"><span data-stu-id="778f7-352">Date</span></span>|<span data-ttu-id="778f7-353">Объект Date</span><span class="sxs-lookup"><span data-stu-id="778f7-353">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778f7-354">Требования</span><span class="sxs-lookup"><span data-stu-id="778f7-354">Requirements</span></span>

|<span data-ttu-id="778f7-355">Требование</span><span class="sxs-lookup"><span data-stu-id="778f7-355">Requirement</span></span>| <span data-ttu-id="778f7-356">Значение</span><span class="sxs-lookup"><span data-stu-id="778f7-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="778f7-357">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="778f7-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778f7-358">1.0</span><span class="sxs-lookup"><span data-stu-id="778f7-358">1.0</span></span>|
|[<span data-ttu-id="778f7-359">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="778f7-359">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778f7-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-360">ReadItem</span></span>|
|[<span data-ttu-id="778f7-361">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="778f7-361">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778f7-362">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-362">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="778f7-363">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="778f7-363">Returns:</span></span>

<span data-ttu-id="778f7-364">Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="778f7-364">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="778f7-365">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="778f7-365">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="778f7-366">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="778f7-366">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="778f7-367">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="778f7-367">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="778f7-p108">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="778f7-p108">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="778f7-370">Параметры</span><span class="sxs-lookup"><span data-stu-id="778f7-370">Parameters</span></span>

|<span data-ttu-id="778f7-371">Имя</span><span class="sxs-lookup"><span data-stu-id="778f7-371">Name</span></span>| <span data-ttu-id="778f7-372">Тип</span><span class="sxs-lookup"><span data-stu-id="778f7-372">Type</span></span>| <span data-ttu-id="778f7-373">Описание</span><span class="sxs-lookup"><span data-stu-id="778f7-373">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="778f7-374">Строка</span><span class="sxs-lookup"><span data-stu-id="778f7-374">String</span></span>|<span data-ttu-id="778f7-375">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="778f7-375">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="778f7-376">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="778f7-376">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="778f7-377">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="778f7-377">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778f7-378">Требования</span><span class="sxs-lookup"><span data-stu-id="778f7-378">Requirements</span></span>

|<span data-ttu-id="778f7-379">Требование</span><span class="sxs-lookup"><span data-stu-id="778f7-379">Requirement</span></span>| <span data-ttu-id="778f7-380">Значение</span><span class="sxs-lookup"><span data-stu-id="778f7-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="778f7-381">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="778f7-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778f7-382">1.3</span><span class="sxs-lookup"><span data-stu-id="778f7-382">1.3</span></span>|
|[<span data-ttu-id="778f7-383">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="778f7-383">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778f7-384">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="778f7-384">Restricted</span></span>|
|[<span data-ttu-id="778f7-385">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="778f7-385">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778f7-386">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-386">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="778f7-387">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="778f7-387">Returns:</span></span>

<span data-ttu-id="778f7-388">Тип: String</span><span class="sxs-lookup"><span data-stu-id="778f7-388">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="778f7-389">Пример</span><span class="sxs-lookup"><span data-stu-id="778f7-389">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="778f7-390">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="778f7-390">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="778f7-391">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="778f7-391">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="778f7-392">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="778f7-392">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="778f7-393">Параметры</span><span class="sxs-lookup"><span data-stu-id="778f7-393">Parameters</span></span>

|<span data-ttu-id="778f7-394">Имя</span><span class="sxs-lookup"><span data-stu-id="778f7-394">Name</span></span>| <span data-ttu-id="778f7-395">Тип</span><span class="sxs-lookup"><span data-stu-id="778f7-395">Type</span></span>| <span data-ttu-id="778f7-396">Описание</span><span class="sxs-lookup"><span data-stu-id="778f7-396">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="778f7-397">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="778f7-397">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="778f7-398">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="778f7-398">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778f7-399">Требования</span><span class="sxs-lookup"><span data-stu-id="778f7-399">Requirements</span></span>

|<span data-ttu-id="778f7-400">Требование</span><span class="sxs-lookup"><span data-stu-id="778f7-400">Requirement</span></span>| <span data-ttu-id="778f7-401">Значение</span><span class="sxs-lookup"><span data-stu-id="778f7-401">Value</span></span>|
|---|---|
|[<span data-ttu-id="778f7-402">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="778f7-402">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778f7-403">1.0</span><span class="sxs-lookup"><span data-stu-id="778f7-403">1.0</span></span>|
|[<span data-ttu-id="778f7-404">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="778f7-404">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778f7-405">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-405">ReadItem</span></span>|
|[<span data-ttu-id="778f7-406">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="778f7-406">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778f7-407">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-407">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="778f7-408">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="778f7-408">Returns:</span></span>

<span data-ttu-id="778f7-409">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="778f7-409">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="778f7-410">Тип: Date</span><span class="sxs-lookup"><span data-stu-id="778f7-410">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="778f7-411">Пример</span><span class="sxs-lookup"><span data-stu-id="778f7-411">Example</span></span>

```js
// Represents 3:37 PM PDT on Monday, August 26, 2019.
var input = {
  date: 26,
  hours: 15,
  milliseconds: 2,
  minutes: 37,
  month: 7,
  seconds: 2,
  timezoneOffset: -420,
  year: 2019
};

// result should be a Date object.
var result = Office.context.mailbox.convertToUtcClientTime(input);

// Output should be "2019-08-26T22:37:02.002Z".
console.log(result.toISOString());
```

<br>

---
---

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="778f7-412">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="778f7-412">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="778f7-413">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="778f7-413">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="778f7-414">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="778f7-414">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="778f7-415">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="778f7-415">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="778f7-p109">В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).</span><span class="sxs-lookup"><span data-stu-id="778f7-p109">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="778f7-418">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="778f7-418">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="778f7-419">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="778f7-419">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="778f7-420">Параметры</span><span class="sxs-lookup"><span data-stu-id="778f7-420">Parameters</span></span>

|<span data-ttu-id="778f7-421">Имя</span><span class="sxs-lookup"><span data-stu-id="778f7-421">Name</span></span>| <span data-ttu-id="778f7-422">Тип</span><span class="sxs-lookup"><span data-stu-id="778f7-422">Type</span></span>| <span data-ttu-id="778f7-423">Описание</span><span class="sxs-lookup"><span data-stu-id="778f7-423">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="778f7-424">String</span><span class="sxs-lookup"><span data-stu-id="778f7-424">String</span></span>|<span data-ttu-id="778f7-425">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="778f7-425">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778f7-426">Требования</span><span class="sxs-lookup"><span data-stu-id="778f7-426">Requirements</span></span>

|<span data-ttu-id="778f7-427">Требование</span><span class="sxs-lookup"><span data-stu-id="778f7-427">Requirement</span></span>| <span data-ttu-id="778f7-428">Значение</span><span class="sxs-lookup"><span data-stu-id="778f7-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="778f7-429">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="778f7-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778f7-430">1.0</span><span class="sxs-lookup"><span data-stu-id="778f7-430">1.0</span></span>|
|[<span data-ttu-id="778f7-431">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="778f7-431">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778f7-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-432">ReadItem</span></span>|
|[<span data-ttu-id="778f7-433">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="778f7-433">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778f7-434">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-434">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="778f7-435">Пример</span><span class="sxs-lookup"><span data-stu-id="778f7-435">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="778f7-436">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="778f7-436">displayMessageForm(itemId)</span></span>

<span data-ttu-id="778f7-437">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="778f7-437">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="778f7-438">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="778f7-438">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="778f7-439">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="778f7-439">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="778f7-440">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="778f7-440">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="778f7-441">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="778f7-441">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="778f7-p110">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="778f7-p110">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="778f7-444">Параметры</span><span class="sxs-lookup"><span data-stu-id="778f7-444">Parameters</span></span>

|<span data-ttu-id="778f7-445">Имя</span><span class="sxs-lookup"><span data-stu-id="778f7-445">Name</span></span>| <span data-ttu-id="778f7-446">Тип</span><span class="sxs-lookup"><span data-stu-id="778f7-446">Type</span></span>| <span data-ttu-id="778f7-447">Описание</span><span class="sxs-lookup"><span data-stu-id="778f7-447">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="778f7-448">Строка</span><span class="sxs-lookup"><span data-stu-id="778f7-448">String</span></span>|<span data-ttu-id="778f7-449">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="778f7-449">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778f7-450">Требования</span><span class="sxs-lookup"><span data-stu-id="778f7-450">Requirements</span></span>

|<span data-ttu-id="778f7-451">Требование</span><span class="sxs-lookup"><span data-stu-id="778f7-451">Requirement</span></span>| <span data-ttu-id="778f7-452">Значение</span><span class="sxs-lookup"><span data-stu-id="778f7-452">Value</span></span>|
|---|---|
|[<span data-ttu-id="778f7-453">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="778f7-453">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778f7-454">1.0</span><span class="sxs-lookup"><span data-stu-id="778f7-454">1.0</span></span>|
|[<span data-ttu-id="778f7-455">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="778f7-455">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778f7-456">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-456">ReadItem</span></span>|
|[<span data-ttu-id="778f7-457">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="778f7-457">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778f7-458">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-458">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="778f7-459">Пример</span><span class="sxs-lookup"><span data-stu-id="778f7-459">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="778f7-460">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="778f7-460">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="778f7-461">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="778f7-461">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="778f7-462">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="778f7-462">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="778f7-p111">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="778f7-p111">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="778f7-p112">В Outlook в Интернете и на мобильных устройствах этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="778f7-p112">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="778f7-p113">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="778f7-p113">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="778f7-470">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="778f7-470">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="778f7-471">Параметры</span><span class="sxs-lookup"><span data-stu-id="778f7-471">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="778f7-472">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="778f7-472">All parameters are optional.</span></span>

|<span data-ttu-id="778f7-473">Имя</span><span class="sxs-lookup"><span data-stu-id="778f7-473">Name</span></span>| <span data-ttu-id="778f7-474">Тип</span><span class="sxs-lookup"><span data-stu-id="778f7-474">Type</span></span>| <span data-ttu-id="778f7-475">Описание</span><span class="sxs-lookup"><span data-stu-id="778f7-475">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="778f7-476">Object</span><span class="sxs-lookup"><span data-stu-id="778f7-476">Object</span></span> | <span data-ttu-id="778f7-477">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="778f7-477">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="778f7-478">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="778f7-478">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="778f7-p114">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="778f7-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="778f7-481">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="778f7-481">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="778f7-p115">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="778f7-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="778f7-484">Date</span><span class="sxs-lookup"><span data-stu-id="778f7-484">Date</span></span> | <span data-ttu-id="778f7-485">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="778f7-485">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="778f7-486">Date</span><span class="sxs-lookup"><span data-stu-id="778f7-486">Date</span></span> | <span data-ttu-id="778f7-487">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="778f7-487">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="778f7-488">Строка</span><span class="sxs-lookup"><span data-stu-id="778f7-488">String</span></span> | <span data-ttu-id="778f7-p116">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="778f7-p116">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="778f7-491">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="778f7-491">Array.&lt;String&gt;</span></span> | <span data-ttu-id="778f7-p117">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="778f7-p117">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="778f7-494">String</span><span class="sxs-lookup"><span data-stu-id="778f7-494">String</span></span> | <span data-ttu-id="778f7-p118">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="778f7-p118">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="778f7-497">String</span><span class="sxs-lookup"><span data-stu-id="778f7-497">String</span></span> | <span data-ttu-id="778f7-p119">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="778f7-p119">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="778f7-500">Требования</span><span class="sxs-lookup"><span data-stu-id="778f7-500">Requirements</span></span>

|<span data-ttu-id="778f7-501">Требование</span><span class="sxs-lookup"><span data-stu-id="778f7-501">Requirement</span></span>| <span data-ttu-id="778f7-502">Значение</span><span class="sxs-lookup"><span data-stu-id="778f7-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="778f7-503">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="778f7-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778f7-504">1.0</span><span class="sxs-lookup"><span data-stu-id="778f7-504">1.0</span></span>|
|[<span data-ttu-id="778f7-505">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="778f7-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778f7-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-506">ReadItem</span></span>|
|[<span data-ttu-id="778f7-507">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="778f7-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778f7-508">Чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-508">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="778f7-509">Пример</span><span class="sxs-lookup"><span data-stu-id="778f7-509">Example</span></span>

```js
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

<br>

---
---

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="778f7-510">Дисплайневмессажеформ (Parameters)</span><span class="sxs-lookup"><span data-stu-id="778f7-510">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="778f7-511">Отображает форму для создания нового сообщения.</span><span class="sxs-lookup"><span data-stu-id="778f7-511">Displays a form for creating a new message.</span></span>

<span data-ttu-id="778f7-512">`displayNewMessageForm` Метод открывает форму, которая позволяет пользователю создать новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="778f7-512">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="778f7-513">Если указаны параметры, поля формы сообщения автоматически заполняются содержимым параметров.</span><span class="sxs-lookup"><span data-stu-id="778f7-513">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="778f7-514">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="778f7-514">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="778f7-515">Параметры</span><span class="sxs-lookup"><span data-stu-id="778f7-515">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="778f7-516">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="778f7-516">All parameters are optional.</span></span>

|<span data-ttu-id="778f7-517">Имя</span><span class="sxs-lookup"><span data-stu-id="778f7-517">Name</span></span>| <span data-ttu-id="778f7-518">Тип</span><span class="sxs-lookup"><span data-stu-id="778f7-518">Type</span></span>| <span data-ttu-id="778f7-519">Описание</span><span class="sxs-lookup"><span data-stu-id="778f7-519">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="778f7-520">Object</span><span class="sxs-lookup"><span data-stu-id="778f7-520">Object</span></span> | <span data-ttu-id="778f7-521">Словарь параметров, описывающих новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="778f7-521">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="778f7-522">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="778f7-522">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="778f7-523">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого из получателей в строке "Кому".</span><span class="sxs-lookup"><span data-stu-id="778f7-523">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="778f7-524">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="778f7-524">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="778f7-525">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="778f7-525">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="778f7-526">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого получателя в строке "копия".</span><span class="sxs-lookup"><span data-stu-id="778f7-526">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="778f7-527">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="778f7-527">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="778f7-528">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="778f7-528">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="778f7-529">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого из получателей, указанных в строке "СК".</span><span class="sxs-lookup"><span data-stu-id="778f7-529">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="778f7-530">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="778f7-530">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="778f7-531">Строка</span><span class="sxs-lookup"><span data-stu-id="778f7-531">String</span></span> | <span data-ttu-id="778f7-532">Строка, содержащая тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="778f7-532">A string containing the subject of the message.</span></span> <span data-ttu-id="778f7-533">Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="778f7-533">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="778f7-534">String</span><span class="sxs-lookup"><span data-stu-id="778f7-534">String</span></span> | <span data-ttu-id="778f7-535">Текст сообщения в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="778f7-535">The HTML body of the message.</span></span> <span data-ttu-id="778f7-536">Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="778f7-536">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="778f7-537">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="778f7-537">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="778f7-538">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="778f7-538">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="778f7-539">Строка</span><span class="sxs-lookup"><span data-stu-id="778f7-539">String</span></span> | <span data-ttu-id="778f7-p126">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="778f7-p126">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="778f7-542">Строка</span><span class="sxs-lookup"><span data-stu-id="778f7-542">String</span></span> | <span data-ttu-id="778f7-543">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="778f7-543">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="778f7-544">Строка</span><span class="sxs-lookup"><span data-stu-id="778f7-544">String</span></span> | <span data-ttu-id="778f7-p127">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="778f7-p127">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="778f7-547">Логический</span><span class="sxs-lookup"><span data-stu-id="778f7-547">Boolean</span></span> | <span data-ttu-id="778f7-p128">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="778f7-p128">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="778f7-550">Строка</span><span class="sxs-lookup"><span data-stu-id="778f7-550">String</span></span> | <span data-ttu-id="778f7-551">Используется, только если свойству `type` присвоено значение `item`.</span><span class="sxs-lookup"><span data-stu-id="778f7-551">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="778f7-552">Идентификатор элемента EWS существующего сообщения электронной почты, которое необходимо присоединить к новому сообщению.</span><span class="sxs-lookup"><span data-stu-id="778f7-552">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="778f7-553">Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="778f7-553">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="778f7-554">Требования</span><span class="sxs-lookup"><span data-stu-id="778f7-554">Requirements</span></span>

|<span data-ttu-id="778f7-555">Требование</span><span class="sxs-lookup"><span data-stu-id="778f7-555">Requirement</span></span>| <span data-ttu-id="778f7-556">Значение</span><span class="sxs-lookup"><span data-stu-id="778f7-556">Value</span></span>|
|---|---|
|[<span data-ttu-id="778f7-557">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="778f7-557">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778f7-558">1.6</span><span class="sxs-lookup"><span data-stu-id="778f7-558">1.6</span></span> |
|[<span data-ttu-id="778f7-559">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="778f7-559">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778f7-560">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-560">ReadItem</span></span>|
|[<span data-ttu-id="778f7-561">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="778f7-561">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778f7-562">Чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-562">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="778f7-563">Пример</span><span class="sxs-lookup"><span data-stu-id="778f7-563">Example</span></span>

```js
Office.context.mailbox.displayNewMessageForm(
  {
    // Copy the To line from current item.
    toRecipients: Office.context.mailbox.item.to,
    ccRecipients: ['sam@contoso.com'],
    subject: 'Outlook add-ins are cool!',
    htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
    attachments: [
      {
        type: 'file',
        name: 'image.png',
        url: 'http://contoso.com/image.png',
        isInline: true
      }
    ]
  });
```

<br>

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="778f7-564">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="778f7-564">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="778f7-565">Возвращает строку, содержащую маркер, который используется для вызова интерфейсов REST API или веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="778f7-565">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="778f7-p130">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="778f7-p130">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="778f7-568">Рекомендуем сделать так, чтобы по мере возможности надстройки использовали интерфейсы REST API, а не веб-службы Exchange.</span><span class="sxs-lookup"><span data-stu-id="778f7-568">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="778f7-569">Для вызова метода `getCallbackTokenAsync` в режиме чтения требуется минимальный уровень разрешения **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="778f7-569">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="778f7-570">Для вызова `getCallbackTokenAsync` в режиме создания сообщения требуется сохранить элемент.</span><span class="sxs-lookup"><span data-stu-id="778f7-570">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="778f7-571">Для метода [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) требуется минимальный уровень разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="778f7-571">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="778f7-572">**Маркеры REST**</span><span class="sxs-lookup"><span data-stu-id="778f7-572">**REST Tokens**</span></span>

<span data-ttu-id="778f7-p132">Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.</span><span class="sxs-lookup"><span data-stu-id="778f7-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="778f7-576">С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.</span><span class="sxs-lookup"><span data-stu-id="778f7-576">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="778f7-577">**Маркеры EWS**</span><span class="sxs-lookup"><span data-stu-id="778f7-577">**EWS Tokens**</span></span>

<span data-ttu-id="778f7-p133">Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.</span><span class="sxs-lookup"><span data-stu-id="778f7-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="778f7-580">С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.</span><span class="sxs-lookup"><span data-stu-id="778f7-580">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="778f7-581">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="778f7-581">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="778f7-582">Третья система использует маркер в качестве маркера авторизации носителя, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) [или GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange (EWS) для получения вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="778f7-582">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to retrieve an attachment or item.</span></span> <span data-ttu-id="778f7-583">Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="778f7-583">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="778f7-584">Параметры</span><span class="sxs-lookup"><span data-stu-id="778f7-584">Parameters</span></span>

|<span data-ttu-id="778f7-585">Имя</span><span class="sxs-lookup"><span data-stu-id="778f7-585">Name</span></span>| <span data-ttu-id="778f7-586">Тип</span><span class="sxs-lookup"><span data-stu-id="778f7-586">Type</span></span>| <span data-ttu-id="778f7-587">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="778f7-587">Attributes</span></span>| <span data-ttu-id="778f7-588">Описание</span><span class="sxs-lookup"><span data-stu-id="778f7-588">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="778f7-589">Object</span><span class="sxs-lookup"><span data-stu-id="778f7-589">Object</span></span> | <span data-ttu-id="778f7-590">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="778f7-590">&lt;optional&gt;</span></span> | <span data-ttu-id="778f7-591">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="778f7-591">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="778f7-592">Boolean</span><span class="sxs-lookup"><span data-stu-id="778f7-592">Boolean</span></span> |  <span data-ttu-id="778f7-593">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="778f7-593">&lt;optional&gt;</span></span> | <span data-ttu-id="778f7-p135">Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="778f7-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="778f7-596">Object</span><span class="sxs-lookup"><span data-stu-id="778f7-596">Object</span></span> |  <span data-ttu-id="778f7-597">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="778f7-597">&lt;optional&gt;</span></span> | <span data-ttu-id="778f7-598">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="778f7-598">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="778f7-599">функция</span><span class="sxs-lookup"><span data-stu-id="778f7-599">function</span></span>||<span data-ttu-id="778f7-600">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="778f7-600">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="778f7-601">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="778f7-601">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="778f7-602">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="778f7-602">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="778f7-603">Ошибки</span><span class="sxs-lookup"><span data-stu-id="778f7-603">Errors</span></span>

|<span data-ttu-id="778f7-604">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="778f7-604">Error code</span></span>|<span data-ttu-id="778f7-605">Описание</span><span class="sxs-lookup"><span data-stu-id="778f7-605">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="778f7-606">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="778f7-606">The request has failed.</span></span> <span data-ttu-id="778f7-607">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="778f7-607">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="778f7-608">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="778f7-608">The Exchange server returned an error.</span></span> <span data-ttu-id="778f7-609">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="778f7-609">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="778f7-610">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="778f7-610">The user is no longer connected to the network.</span></span> <span data-ttu-id="778f7-611">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="778f7-611">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778f7-612">Требования</span><span class="sxs-lookup"><span data-stu-id="778f7-612">Requirements</span></span>

|<span data-ttu-id="778f7-613">Требование</span><span class="sxs-lookup"><span data-stu-id="778f7-613">Requirement</span></span>| <span data-ttu-id="778f7-614">Значение</span><span class="sxs-lookup"><span data-stu-id="778f7-614">Value</span></span>|
|---|---|
|[<span data-ttu-id="778f7-615">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="778f7-615">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778f7-616">1.5</span><span class="sxs-lookup"><span data-stu-id="778f7-616">1.5</span></span> |
|[<span data-ttu-id="778f7-617">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="778f7-617">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778f7-618">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-618">ReadItem</span></span>|
|[<span data-ttu-id="778f7-619">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="778f7-619">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778f7-620">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-620">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="778f7-621">Пример</span><span class="sxs-lookup"><span data-stu-id="778f7-621">Example</span></span>

```js
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="778f7-622">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="778f7-622">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="778f7-623">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="778f7-623">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="778f7-p139">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="778f7-p139">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="778f7-626">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="778f7-626">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="778f7-627">Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="778f7-627">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="778f7-628">Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="778f7-628">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="778f7-629">Для вызова метода `getCallbackTokenAsync` в режиме чтения требуется минимальный уровень разрешения **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="778f7-629">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="778f7-630">Для вызова `getCallbackTokenAsync` в режиме создания сообщения требуется сохранить элемент.</span><span class="sxs-lookup"><span data-stu-id="778f7-630">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="778f7-631">Для метода [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) требуется минимальный уровень разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="778f7-631">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="778f7-632">Параметры</span><span class="sxs-lookup"><span data-stu-id="778f7-632">Parameters</span></span>

|<span data-ttu-id="778f7-633">Имя</span><span class="sxs-lookup"><span data-stu-id="778f7-633">Name</span></span>| <span data-ttu-id="778f7-634">Тип</span><span class="sxs-lookup"><span data-stu-id="778f7-634">Type</span></span>| <span data-ttu-id="778f7-635">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="778f7-635">Attributes</span></span>| <span data-ttu-id="778f7-636">Описание</span><span class="sxs-lookup"><span data-stu-id="778f7-636">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="778f7-637">функция</span><span class="sxs-lookup"><span data-stu-id="778f7-637">function</span></span>||<span data-ttu-id="778f7-638">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="778f7-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="778f7-639">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="778f7-639">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="778f7-640">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="778f7-640">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="778f7-641">Объект</span><span class="sxs-lookup"><span data-stu-id="778f7-641">Object</span></span>| <span data-ttu-id="778f7-642">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="778f7-642">&lt;optional&gt;</span></span>|<span data-ttu-id="778f7-643">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="778f7-643">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="778f7-644">Ошибки</span><span class="sxs-lookup"><span data-stu-id="778f7-644">Errors</span></span>

|<span data-ttu-id="778f7-645">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="778f7-645">Error code</span></span>|<span data-ttu-id="778f7-646">Описание</span><span class="sxs-lookup"><span data-stu-id="778f7-646">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="778f7-647">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="778f7-647">The request has failed.</span></span> <span data-ttu-id="778f7-648">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="778f7-648">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="778f7-649">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="778f7-649">The Exchange server returned an error.</span></span> <span data-ttu-id="778f7-650">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="778f7-650">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="778f7-651">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="778f7-651">The user is no longer connected to the network.</span></span> <span data-ttu-id="778f7-652">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="778f7-652">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778f7-653">Требования</span><span class="sxs-lookup"><span data-stu-id="778f7-653">Requirements</span></span>

|<span data-ttu-id="778f7-654">Требование</span><span class="sxs-lookup"><span data-stu-id="778f7-654">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="778f7-655">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="778f7-655">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778f7-656">1.0</span><span class="sxs-lookup"><span data-stu-id="778f7-656">1.0</span></span> | <span data-ttu-id="778f7-657">1.3</span><span class="sxs-lookup"><span data-stu-id="778f7-657">1.3</span></span> |
|[<span data-ttu-id="778f7-658">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="778f7-658">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778f7-659">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-659">ReadItem</span></span> | <span data-ttu-id="778f7-660">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-660">ReadItem</span></span> |
|[<span data-ttu-id="778f7-661">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="778f7-661">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778f7-662">Чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-662">Read</span></span> | <span data-ttu-id="778f7-663">Создание</span><span class="sxs-lookup"><span data-stu-id="778f7-663">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="778f7-664">Пример</span><span class="sxs-lookup"><span data-stu-id="778f7-664">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="778f7-665">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="778f7-665">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="778f7-666">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="778f7-666">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="778f7-667">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="778f7-667">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="778f7-668">Параметры</span><span class="sxs-lookup"><span data-stu-id="778f7-668">Parameters</span></span>

|<span data-ttu-id="778f7-669">Имя</span><span class="sxs-lookup"><span data-stu-id="778f7-669">Name</span></span>| <span data-ttu-id="778f7-670">Тип</span><span class="sxs-lookup"><span data-stu-id="778f7-670">Type</span></span>| <span data-ttu-id="778f7-671">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="778f7-671">Attributes</span></span>| <span data-ttu-id="778f7-672">Описание</span><span class="sxs-lookup"><span data-stu-id="778f7-672">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="778f7-673">функция</span><span class="sxs-lookup"><span data-stu-id="778f7-673">function</span></span>||<span data-ttu-id="778f7-674">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="778f7-674">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="778f7-675">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="778f7-675">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="778f7-676">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="778f7-676">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="778f7-677">Объект</span><span class="sxs-lookup"><span data-stu-id="778f7-677">Object</span></span>| <span data-ttu-id="778f7-678">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="778f7-678">&lt;optional&gt;</span></span>|<span data-ttu-id="778f7-679">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="778f7-679">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="778f7-680">Ошибки</span><span class="sxs-lookup"><span data-stu-id="778f7-680">Errors</span></span>

|<span data-ttu-id="778f7-681">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="778f7-681">Error code</span></span>|<span data-ttu-id="778f7-682">Описание</span><span class="sxs-lookup"><span data-stu-id="778f7-682">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="778f7-683">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="778f7-683">The request has failed.</span></span> <span data-ttu-id="778f7-684">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="778f7-684">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="778f7-685">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="778f7-685">The Exchange server returned an error.</span></span> <span data-ttu-id="778f7-686">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="778f7-686">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="778f7-687">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="778f7-687">The user is no longer connected to the network.</span></span> <span data-ttu-id="778f7-688">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="778f7-688">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778f7-689">Требования</span><span class="sxs-lookup"><span data-stu-id="778f7-689">Requirements</span></span>

|<span data-ttu-id="778f7-690">Требование</span><span class="sxs-lookup"><span data-stu-id="778f7-690">Requirement</span></span>| <span data-ttu-id="778f7-691">Значение</span><span class="sxs-lookup"><span data-stu-id="778f7-691">Value</span></span>|
|---|---|
|[<span data-ttu-id="778f7-692">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="778f7-692">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778f7-693">1.0</span><span class="sxs-lookup"><span data-stu-id="778f7-693">1.0</span></span>|
|[<span data-ttu-id="778f7-694">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="778f7-694">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778f7-695">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-695">ReadItem</span></span>|
|[<span data-ttu-id="778f7-696">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="778f7-696">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778f7-697">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-697">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="778f7-698">Пример</span><span class="sxs-lookup"><span data-stu-id="778f7-698">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="778f7-699">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="778f7-699">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="778f7-700">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="778f7-700">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="778f7-701">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="778f7-701">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="778f7-702">В Outlook для iOS и Android</span><span class="sxs-lookup"><span data-stu-id="778f7-702">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="778f7-703">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="778f7-703">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="778f7-704">В таких случаях надстройка должна [использовать REST API](/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="778f7-704">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="778f7-705">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="778f7-705">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="778f7-706">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="778f7-706">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="778f7-707">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="778f7-707">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="778f7-708">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="778f7-708">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="778f7-p149">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="778f7-p149">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="778f7-711">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="778f7-711">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="778f7-712">Различия версий</span><span class="sxs-lookup"><span data-stu-id="778f7-712">Version differences</span></span>

<span data-ttu-id="778f7-713">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="778f7-713">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="778f7-714">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="778f7-714">You do not need to set the encoding value when your mail app is running in Outlook on the web.</span></span> <span data-ttu-id="778f7-715">Вы можете определить, работает ли почтовое приложение в Outlook в Интернете или на настольном клиенте с помощью свойства Mailbox. Diagnostics. hostName.</span><span class="sxs-lookup"><span data-stu-id="778f7-715">You can determine whether your mail app is running in Outlook on the web or a desktop client by using the mailbox.diagnostics.hostName property.</span></span> <span data-ttu-id="778f7-716">Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="778f7-716">You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="778f7-717">Параметры</span><span class="sxs-lookup"><span data-stu-id="778f7-717">Parameters</span></span>

|<span data-ttu-id="778f7-718">Имя</span><span class="sxs-lookup"><span data-stu-id="778f7-718">Name</span></span>| <span data-ttu-id="778f7-719">Тип</span><span class="sxs-lookup"><span data-stu-id="778f7-719">Type</span></span>| <span data-ttu-id="778f7-720">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="778f7-720">Attributes</span></span>| <span data-ttu-id="778f7-721">Описание</span><span class="sxs-lookup"><span data-stu-id="778f7-721">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="778f7-722">String</span><span class="sxs-lookup"><span data-stu-id="778f7-722">String</span></span>||<span data-ttu-id="778f7-723">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="778f7-723">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="778f7-724">function</span><span class="sxs-lookup"><span data-stu-id="778f7-724">function</span></span>||<span data-ttu-id="778f7-725">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="778f7-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="778f7-726">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="778f7-726">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="778f7-727">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="778f7-727">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="778f7-728">Object</span><span class="sxs-lookup"><span data-stu-id="778f7-728">Object</span></span>| <span data-ttu-id="778f7-729">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="778f7-729">&lt;optional&gt;</span></span>|<span data-ttu-id="778f7-730">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="778f7-730">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778f7-731">Требования</span><span class="sxs-lookup"><span data-stu-id="778f7-731">Requirements</span></span>

|<span data-ttu-id="778f7-732">Требование</span><span class="sxs-lookup"><span data-stu-id="778f7-732">Requirement</span></span>| <span data-ttu-id="778f7-733">Значение</span><span class="sxs-lookup"><span data-stu-id="778f7-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="778f7-734">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="778f7-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778f7-735">1.0</span><span class="sxs-lookup"><span data-stu-id="778f7-735">1.0</span></span>|
|[<span data-ttu-id="778f7-736">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="778f7-736">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778f7-737">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="778f7-737">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="778f7-738">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="778f7-738">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778f7-739">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-739">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="778f7-740">Пример</span><span class="sxs-lookup"><span data-stu-id="778f7-740">Example</span></span>

<span data-ttu-id="778f7-741">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="778f7-741">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```js
function getSubjectRequest(id) {
  // Return a GetItem operation request for the subject of the specified item.
  var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

  return request;
}

function sendRequest() {
  // Create a local variable that contains the mailbox.
  Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
  var result = asyncResult.value;
  var context = asyncResult.asyncContext;

  // Process the returned response here.
}
```

<br>

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="778f7-742">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="778f7-742">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="778f7-743">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="778f7-743">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="778f7-744">В настоящее время поддерживаются типы событий `Office.EventType.ItemChanged` и `Office.EventType.OfficeThemeChanged`.</span><span class="sxs-lookup"><span data-stu-id="778f7-744">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="778f7-745">Параметры</span><span class="sxs-lookup"><span data-stu-id="778f7-745">Parameters</span></span>

| <span data-ttu-id="778f7-746">Имя</span><span class="sxs-lookup"><span data-stu-id="778f7-746">Name</span></span> | <span data-ttu-id="778f7-747">Тип</span><span class="sxs-lookup"><span data-stu-id="778f7-747">Type</span></span> | <span data-ttu-id="778f7-748">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="778f7-748">Attributes</span></span> | <span data-ttu-id="778f7-749">Описание</span><span class="sxs-lookup"><span data-stu-id="778f7-749">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="778f7-750">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="778f7-750">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="778f7-751">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="778f7-751">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="778f7-752">Object</span><span class="sxs-lookup"><span data-stu-id="778f7-752">Object</span></span> | <span data-ttu-id="778f7-753">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="778f7-753">&lt;optional&gt;</span></span> | <span data-ttu-id="778f7-754">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="778f7-754">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="778f7-755">Object</span><span class="sxs-lookup"><span data-stu-id="778f7-755">Object</span></span> | <span data-ttu-id="778f7-756">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="778f7-756">&lt;optional&gt;</span></span> | <span data-ttu-id="778f7-757">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="778f7-757">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="778f7-758">функция</span><span class="sxs-lookup"><span data-stu-id="778f7-758">function</span></span>| <span data-ttu-id="778f7-759">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="778f7-759">&lt;optional&gt;</span></span>|<span data-ttu-id="778f7-760">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="778f7-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778f7-761">Требования</span><span class="sxs-lookup"><span data-stu-id="778f7-761">Requirements</span></span>

|<span data-ttu-id="778f7-762">Требование</span><span class="sxs-lookup"><span data-stu-id="778f7-762">Requirement</span></span>| <span data-ttu-id="778f7-763">Значение</span><span class="sxs-lookup"><span data-stu-id="778f7-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="778f7-764">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="778f7-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778f7-765">1.5</span><span class="sxs-lookup"><span data-stu-id="778f7-765">1.5</span></span> |
|[<span data-ttu-id="778f7-766">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="778f7-766">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778f7-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778f7-767">ReadItem</span></span> |
|[<span data-ttu-id="778f7-768">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="778f7-768">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778f7-769">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="778f7-769">Compose or Read</span></span>|
