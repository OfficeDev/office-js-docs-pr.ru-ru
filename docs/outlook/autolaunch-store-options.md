---
title: Параметры списка AppSource для надстройки на Outlook событий
description: Узнайте о параметрах списка AppSource, доступных для Outlook надстройки, которая реализует активацию на основе событий.
ms.topic: article
ms.date: 07/14/2021
localization_priority: Normal
ms.openlocfilehash: 0704b96b51841ec70aaf014924bed931c177b467
ms.sourcegitcommit: 30a861ece18255e342725e31c47f01960b854532
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/16/2021
ms.locfileid: "53458951"
---
# <a name="appsource-listing-options-for-your-event-based-outlook-add-in"></a><span data-ttu-id="28163-103">Параметры списка AppSource для надстройки на Outlook событий</span><span class="sxs-lookup"><span data-stu-id="28163-103">AppSource listing options for your event-based Outlook add-in</span></span>

<span data-ttu-id="28163-104">В настоящее время надстройки должны быть развернуты администраторами организации для конечных пользователей, чтобы получить доступ к функциональным возможностям функций на основе событий.</span><span class="sxs-lookup"><span data-stu-id="28163-104">At present, add-ins must be deployed by an organization's admins for end-users to access the event-based feature functionality.</span></span> <span data-ttu-id="28163-105">Мы ограничим активацию на основе событий, если конечный пользователь приобрел надстройку непосредственно в AppSource.</span><span class="sxs-lookup"><span data-stu-id="28163-105">We're restricting event-based activation if the end-user acquired the add-in directly from AppSource.</span></span> <span data-ttu-id="28163-106">Например, если надстройка Contoso включает точку расширения с по крайней мере одной, определенной под узлом (см. следующий отрывок из примера манифеста надстройки), автоматическое вызов надстройки происходит только в том случае, если надстройка была установлена для конечного пользователя администратором организации. В противном случае автоматическое вызов надстройки `LaunchEvent` `LaunchEvent Type` `LaunchEvents` блокируется.</span><span class="sxs-lookup"><span data-stu-id="28163-106">For example, if the Contoso add-in includes the `LaunchEvent` extension point with at least one defined `LaunchEvent Type` under the `LaunchEvents` node (see the following excerpt from an example add-in manifest), the automatic invocation of the add-in only happens if the add-in was installed for the end-user by their organization's admin. Otherwise, the automatic invocation of the add-in is blocked.</span></span>

```xml
...
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    ...
```

<span data-ttu-id="28163-107">Конечный пользователь или администратор могут приобретать надстройки через AppSource или inclient store.</span><span class="sxs-lookup"><span data-stu-id="28163-107">An end-user or admin can acquire add-ins through AppSource or the inclient store.</span></span> <span data-ttu-id="28163-108">Если основной сценарий или рабочий процесс надстройки требует активации на основе событий, возможно, вам потребуется ограничить свои надстройки, доступные для развертывания администратора.</span><span class="sxs-lookup"><span data-stu-id="28163-108">If your add-in's primary scenario or workflow requires event-based activation, you may want to restrict your add-ins available to admin deployment.</span></span> <span data-ttu-id="28163-109">Чтобы включить это ограничение, мы можем предоставить URL-адреса кода полета.</span><span class="sxs-lookup"><span data-stu-id="28163-109">To enable that restriction, we can provide flight code URLs.</span></span> <span data-ttu-id="28163-110">Благодаря кодам полетов доступ к списку могут получить только конечные пользователи с этими специальными URL-адресами.</span><span class="sxs-lookup"><span data-stu-id="28163-110">Thanks to the flight codes, only end-users with these special URLs can access the listing.</span></span> <span data-ttu-id="28163-111">Ниже приводится пример URL-адреса.</span><span class="sxs-lookup"><span data-stu-id="28163-111">The following is an example URL.</span></span>

`https://appsource.microsoft.com/product/office/WA200002862?flightCodes=EventBasedTest1`

<span data-ttu-id="28163-112">Пользователи и администраторы не могут явно искать надстройку по ее имени в AppSource или inclient store, если для нее включен код полета.</span><span class="sxs-lookup"><span data-stu-id="28163-112">Users and admins can't explicitly search for an add-in by its name in AppSource or the inclient store when a flight code is enabled for it.</span></span> <span data-ttu-id="28163-113">Как создатель надстройки, вы можете в частном порядке обмениваться этими кодами полетов с администраторами организации для развертывания надстройки.</span><span class="sxs-lookup"><span data-stu-id="28163-113">As the add-in creator, you can privately share these flight codes with organization admins for add-in deployment.</span></span>

> [!NOTE]
> <span data-ttu-id="28163-114">Хотя конечные пользователи могут установить надстройки с помощью кода полета, надстройка не будет включать активацию на основе событий.</span><span class="sxs-lookup"><span data-stu-id="28163-114">While end-users can install the add-in using a flight code, the add-in won't include event-based activation.</span></span>

## <a name="specify-a-flight-code"></a><span data-ttu-id="28163-115">Указание кода полета</span><span class="sxs-lookup"><span data-stu-id="28163-115">Specify a flight code</span></span>

<span data-ttu-id="28163-116">Вы можете указать код полета для надстройки,  поделившись этой информацией в примечаниях для сертификации при публикации надстройки.</span><span class="sxs-lookup"><span data-stu-id="28163-116">You can specify the flight code you want for your add-in by sharing that information in the **Notes for certification** when you're publishing your add-in.</span></span> <span data-ttu-id="28163-117">_**Важно:**_ Коды полетов являются конфиденциальными.</span><span class="sxs-lookup"><span data-stu-id="28163-117">_**Important**:_ Flight codes are case-sensitive.</span></span>

![Снимок экрана, показывающий пример запроса кода полета в Примечаниях для экрана сертификации во время публикации.](../images/outlook-publish-notes-for-certification-1.png)

## <a name="deploy-add-in-with-flight-code"></a><span data-ttu-id="28163-119">Развертывание надстройки с кодом полета</span><span class="sxs-lookup"><span data-stu-id="28163-119">Deploy add-in with flight code</span></span>

<span data-ttu-id="28163-120">После задав коды полетов, вы получите URL-адрес из группы сертификации приложений.</span><span class="sxs-lookup"><span data-stu-id="28163-120">After the flight codes are set, you'll receive the URL from the app certification team.</span></span> <span data-ttu-id="28163-121">Затем вы можете поделиться URL-адресом с администраторами в частном порядке.</span><span class="sxs-lookup"><span data-stu-id="28163-121">You can then share the URL with admins privately.</span></span>

<span data-ttu-id="28163-122">Для развертывания надстройки администратор может использовать следующие действия.</span><span class="sxs-lookup"><span data-stu-id="28163-122">To deploy the add-in, the admin can use the following steps.</span></span>

- <span data-ttu-id="28163-123">Во входе admin.microsoft.com или AppSource.com учетную запись Microsoft 365 администратора.</span><span class="sxs-lookup"><span data-stu-id="28163-123">Sign in to admin.microsoft.com or AppSource.com with your Microsoft 365 admin account.</span></span> <span data-ttu-id="28163-124">Если надстройка включена с одним входом (SSO), необходимы глобальные учетные данные администратора.</span><span class="sxs-lookup"><span data-stu-id="28163-124">If the add-in has Single sign-on (SSO) enabled, global admin credentials are needed.</span></span>
- <span data-ttu-id="28163-125">Откройте URL-адрес кода полета в веб-браузере.</span><span class="sxs-lookup"><span data-stu-id="28163-125">Open the flight code URL into a web browser.</span></span>
- <span data-ttu-id="28163-126">На странице списка надстройки выберите **Get it now**.</span><span class="sxs-lookup"><span data-stu-id="28163-126">On the add-in listing page, select **Get it now**.</span></span> <span data-ttu-id="28163-127">Вы должны быть перенаправлены на портал интегрированных приложений.</span><span class="sxs-lookup"><span data-stu-id="28163-127">You should be redirected to the integrated app portal.</span></span>

## <a name="unrestricted-appsource-listing"></a><span data-ttu-id="28163-128">Неограниченное перечисление AppSource</span><span class="sxs-lookup"><span data-stu-id="28163-128">Unrestricted AppSource listing</span></span>

<span data-ttu-id="28163-129">Если надстройка не использует активацию на основе событий для критических сценариев (то есть надстройка работает хорошо без автоматического вызовов), рассмотрите возможность включения надстройки в AppSource без специальных кодов полетов.</span><span class="sxs-lookup"><span data-stu-id="28163-129">If your add-in doesn't use event-based activation for critical scenarios (that is, your add-in works well without automatic invocation), consider listing your add-in in AppSource without any special flight codes.</span></span> <span data-ttu-id="28163-130">Если конечный пользователь получает надстройки из AppSource, автоматическая активация не произойдет для пользователя.</span><span class="sxs-lookup"><span data-stu-id="28163-130">If an end-user gets your add-in from AppSource, automatic activation won't happen for the user.</span></span> <span data-ttu-id="28163-131">Однако они могут использовать другие компоненты надстройки, такие как области задач или команды без пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="28163-131">However, they can use other components of your add-in such as a task pane or UI-less command.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="28163-132">Это временное ограничение.</span><span class="sxs-lookup"><span data-stu-id="28163-132">This is a temporary restriction.</span></span> <span data-ttu-id="28163-133">В будущем мы планируем включить активацию надстройки на основе событий для конечных пользователей, непосредственно приобретавших надстройки.</span><span class="sxs-lookup"><span data-stu-id="28163-133">In future, we plan to enable event-based add-in activation for end-users who directly acquire your add-in.</span></span>

## <a name="update-existing-add-ins-to-include-event-based-activation"></a><span data-ttu-id="28163-134">Обновление существующих надстройок, чтобы включить активацию на основе событий</span><span class="sxs-lookup"><span data-stu-id="28163-134">Update existing add-ins to include event-based activation</span></span>

<span data-ttu-id="28163-135">Вы можете обновить существующую надстройка, чтобы включить активацию на основе событий, а затем повторно переподключить ее для проверки и решить, хотите ли вы иметь ограниченный или неограниченный список AppSource.</span><span class="sxs-lookup"><span data-stu-id="28163-135">You can update your existing add-in to include event-based activation then resubmit it for validation and decide if you want a restricted or unrestricted AppSource listing.</span></span>

<span data-ttu-id="28163-136">После утверждения обновленной надстройки администраторы организации, которые уже развернули надстройки, получат сообщение об обновлении на портале администрирования.</span><span class="sxs-lookup"><span data-stu-id="28163-136">After the updated add-in has been approved, organization admins who have already deployed the add-in will receive an update message in the admin portal.</span></span> <span data-ttu-id="28163-137">Сообщение сообщает администратору об изменениях активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="28163-137">The message advises the admin about the event-based activation changes.</span></span> <span data-ttu-id="28163-138">После того как администратор примет изменения, обновление будет развернуто для конечных пользователей.</span><span class="sxs-lookup"><span data-stu-id="28163-138">After the admin accepts the changes, the update will be deployed to end-users.</span></span>

![Снимок экрана уведомления об обновлении приложения на экране "Интегрированные приложения".](../images/outlook-deploy-update-notification.png)

<span data-ttu-id="28163-140">Для конечных пользователей, которые самостоятельно установили надстройки, функция активации на основе событий не будет работать даже после обновления надстройки.</span><span class="sxs-lookup"><span data-stu-id="28163-140">For end-users who installed the add-in on their own, the event-based activation feature won't work even after the add-in has been updated.</span></span>

## <a name="admin-consent-for-installing-event-based-add-ins"></a><span data-ttu-id="28163-141">Согласие администратора на установку надстройок на основе событий</span><span class="sxs-lookup"><span data-stu-id="28163-141">Admin consent for installing event-based add-ins</span></span>

<span data-ttu-id="28163-142">Всякий раз, когда надстройка  на основе событий развертывается из раздела Интегрированные приложения центра администрирования, администратор получает сведения о возможностях активации на основе событий надстройки в мастере развертывания.</span><span class="sxs-lookup"><span data-stu-id="28163-142">Whenever an event-based add-in is deployed from the **Integrated Apps** section of the admin center, the admin gets details about the add-in's event-based activation capabilities in the deployment wizard.</span></span> <span data-ttu-id="28163-143">Сведения отображаются в **разделе Разрешения и возможности приложения.**</span><span class="sxs-lookup"><span data-stu-id="28163-143">The details appear in the **App Permissions and Capabilities** section.</span></span> <span data-ttu-id="28163-144">Администратор должен видеть все события, в которых надстройка может автоматически активироваться.</span><span class="sxs-lookup"><span data-stu-id="28163-144">The admin should see all the events where the add-in can automatically activate.</span></span>

![Снимок экрана "Прием запросов разрешений" при развертывании нового приложения.](../images/outlook-deploy-accept-permissions-requests.png)

<span data-ttu-id="28163-146">Аналогичным образом, когда существующая надстройка обновляется до функции на основе событий, администратор видит в надстройки состояние "Обновление в ожидании".</span><span class="sxs-lookup"><span data-stu-id="28163-146">Similarly, when an existing add-in is updated to event-based functionality, the admin sees an "Update Pending" status on the add-in.</span></span> <span data-ttu-id="28163-147">Обновленная надстройка развертывается только в том случае, если  администратор соглашается на изменения, отмеченные в разделе Разрешения и возможности приложения, включая набор событий, в которых надстройка может автоматически активироваться.</span><span class="sxs-lookup"><span data-stu-id="28163-147">The updated add-in is deployed only if the admin consents to the changes noted in the **App Permissions and Capabilities** section, including the set of events where the add-in can automatically activate.</span></span>

<span data-ttu-id="28163-148">Каждый раз, когда вы добавляете какие-либо новые в надстройку, администраторы будут видеть поток обновления на портале администрирования и должны предоставить согласие `LaunchEvent Type` на дополнительные события.</span><span class="sxs-lookup"><span data-stu-id="28163-148">Each time you add any new `LaunchEvent Type` to your add-in, admins will see the update flow in the admin portal and need to provide consent for additional events.</span></span>

![Снимок экрана потока "Обновления" при развертывании обновленного приложения.](../images/outlook-deploy-update-flow.png)

## <a name="see-also"></a><span data-ttu-id="28163-150">См. также</span><span class="sxs-lookup"><span data-stu-id="28163-150">See also</span></span>

- [<span data-ttu-id="28163-151">Настройка надстройки Outlook для активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="28163-151">Configure your Outlook add-in for event-based activation</span></span>](autolaunch.md)
