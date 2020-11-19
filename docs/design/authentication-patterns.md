---
title: Рекомендации по разработке аутентификации для надстроек Office
ms.date: 07/30/2020
description: Сведения о том, как визуально разработать страницу входа или регистрации в надстройке Office.
localization_priority: Normal
ms.openlocfilehash: e75ab26f820e53ca0379467fba559ebcb727ee3a
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132146"
---
# <a name="authentication-patterns"></a><span data-ttu-id="39664-103">Шаблоны аутентификации</span><span class="sxs-lookup"><span data-stu-id="39664-103">Authentication patterns</span></span>

<span data-ttu-id="39664-104">Для получения доступа к функциям надстройки может требоваться вход или регистрация.</span><span class="sxs-lookup"><span data-stu-id="39664-104">Add-ins may require users to sign-in or sign-up in order to access features and functionality.</span></span> <span data-ttu-id="39664-105">В интерфейс часто встраиваются поля для ввода имени пользователя и пароля или кнопки, которые запускают сторонние потоки идентификации.</span><span class="sxs-lookup"><span data-stu-id="39664-105">Input boxes for username and password or buttons that start third party credential flows are common interface controls in authentication experiences.</span></span> <span data-ttu-id="39664-106">Простая и эффективная аутентификация — важный первый шаг к началу работы с надстройкой.</span><span class="sxs-lookup"><span data-stu-id="39664-106">A simple and efficient authentication experience is an important first step to getting users started with your add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="39664-107">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="39664-107">Best practices</span></span>

|<span data-ttu-id="39664-108">Правильно</span><span class="sxs-lookup"><span data-stu-id="39664-108">Do</span></span>|<span data-ttu-id="39664-109">Неправильно</span><span class="sxs-lookup"><span data-stu-id="39664-109">Don't</span></span>|
|:----|:----|
|<span data-ttu-id="39664-110">Опишите значение надстройки или продемонстрируйте функции, не требуя создания учетной записи.</span><span class="sxs-lookup"><span data-stu-id="39664-110">Prior to sign-in, describe the value of your add-in or demonstrate functionality without requiring an account.</span></span> |<span data-ttu-id="39664-111">Не ожидайте, что пользователи выполнят вход, не понимая значения и преимуществ надстройки.</span><span class="sxs-lookup"><span data-stu-id="39664-111">Expect users to sign-in without understanding the value and benefits of your add-in.</span></span>|
|<span data-ttu-id="39664-112">Направляйте пользователей, используя основную, хорошо видимую кнопку на каждом экране.</span><span class="sxs-lookup"><span data-stu-id="39664-112">Guide users through authentication flows with a primary, highly visible button on each screen.</span></span> |<span data-ttu-id="39664-113">Не обращайте внимание на второстепенные и производные задачи с помощью конкурирующих кнопок и призывов к действию.</span><span class="sxs-lookup"><span data-stu-id="39664-113">Draw attention to secondary and tertiary tasks with competing buttons and calls to action.</span></span>|
|<span data-ttu-id="39664-114">Используйте понятные подписи кнопок с указанием конкретных задач, например "Войти" или "Создать учетную запись".</span><span class="sxs-lookup"><span data-stu-id="39664-114">Use clear button labels that describe specific tasks like "Sign in" or "Create account".</span></span> |<span data-ttu-id="39664-115">Не используйте абстрактные подписи, например "Отправить" или "Начать".</span><span class="sxs-lookup"><span data-stu-id="39664-115">Use vague button labels like "Submit" or "Get started" to guide users through authentication flows.</span></span>|
|<span data-ttu-id="39664-116">Используйте диалоговое окно, чтобы обратить внимание пользователей на формы аутентификации.</span><span class="sxs-lookup"><span data-stu-id="39664-116">Use a dialog to focus users' attention on authentication forms.</span></span> |<span data-ttu-id="39664-117">Не перегружайте область задач инструкциями при первом запуске и формами аутентификации.</span><span class="sxs-lookup"><span data-stu-id="39664-117">Overcrowd your task pane with a first run experience and authentication forms.</span></span>|
|<span data-ttu-id="39664-118">Добавьте небольшие полезные действия, например автофокусировку на полях ввода.</span><span class="sxs-lookup"><span data-stu-id="39664-118">Find small efficiencies in the flow like auto-focusing on input boxes.</span></span> |<span data-ttu-id="39664-119">Не добавляйте ненужные шаги, например не требуйте нажимать на поля формы.</span><span class="sxs-lookup"><span data-stu-id="39664-119">Add unnecessary steps to the interaction like requiring users to click into form fields.</span></span>|
|<span data-ttu-id="39664-120">Предоставить пользователям способ выхода и проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="39664-120">Provide a way for users to sign out and reauthenticate.</span></span> |<span data-ttu-id="39664-121">Не заставляйте пользователей удалять надстройку, чтобы сменить учетную запись.</span><span class="sxs-lookup"><span data-stu-id="39664-121">Force users to uninstall to switch identities.</span></span>|

## <a name="authentication-flow"></a><span data-ttu-id="39664-122">Последовательность аутентификации</span><span class="sxs-lookup"><span data-stu-id="39664-122">Authentication flow</span></span>

1. <span data-ttu-id="39664-123">Первый запуск. Разместите кнопку для входа как четкий призыв к действию при первом запуске надстройки.</span><span class="sxs-lookup"><span data-stu-id="39664-123">First Run Placemat - Place your sign-in button as a clear call-to action inside your add-in's first run experience.</span></span>

    ![Снимок экрана: область задач надстройки в приложении Office](../images/add-in-fre-value-placemat.png)

1. <span data-ttu-id="39664-125">Диалоговое окно выбора службы идентификации. Покажите список служб идентификации, включая, при необходимости, форму для ввода имени пользователя и пароля.</span><span class="sxs-lookup"><span data-stu-id="39664-125">Identity Provider Choices Dialog - Display a clear list of identity providers including a username and password form if applicable.</span></span> <span data-ttu-id="39664-126">Пользовательский интерфейс вашей надстройки может быть заблокирован, когда открыто диалоговое окно аутентификации.</span><span class="sxs-lookup"><span data-stu-id="39664-126">Your add-in UI may be blocked while the authentication dialog is open.</span></span>

    ![Снимок экрана: диалоговое окно выбора поставщика удостоверений в приложении Office](../images/add-in-auth-choices-dialog.png)

1. <span data-ttu-id="39664-128">Вход через службу идентификации. Отобразится пользовательский интерфейс службы идентификации.</span><span class="sxs-lookup"><span data-stu-id="39664-128">Identity Provider Sign-in - The identity provider will have their own UI.</span></span> <span data-ttu-id="39664-129">Microsoft Azure Active Directory позволяет настраивать страницы входа и панели доступа для согласованного внешнего вида со службой.</span><span class="sxs-lookup"><span data-stu-id="39664-129">Microsoft Azure Active Directory allows customization of sign-in and access panel pages for consistent look and feel with your service.</span></span> <span data-ttu-id="39664-130">Дополнительные [сведения](/azure/active-directory/fundamentals/customize-branding).</span><span class="sxs-lookup"><span data-stu-id="39664-130">[Learn More](/azure/active-directory/fundamentals/customize-branding).</span></span>

    ![Снимок экрана с диалоговым окном входа поставщика удостоверений в приложении Office](../images/add-in-auth-identity-sign-in.png)

1. <span data-ttu-id="39664-132">Ход выполнения. Показывайте ход загрузки параметров и пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="39664-132">Progress - Indicate progress while settings and UI load.</span></span>

    ![Снимок экрана с диалоговым окном с индикатором хода выполнения в приложении Office](../images/add-in-auth-modal-interstitial.png)

> [!NOTE]
> <span data-ttu-id="39664-134">Используя службу идентификации Майкрософт, вы получите возможность использовать фирменную кнопку входа, которую можно настроить под светлую и темную темы.</span><span class="sxs-lookup"><span data-stu-id="39664-134">When using Microsoft's Identity service you'll have the opportunity to use a branded sign-in button that is customizable to light and dark themes.</span></span> <span data-ttu-id="39664-135">Узнайте больше.</span><span class="sxs-lookup"><span data-stu-id="39664-135">Learn more.</span></span>

## <a name="single-sign-on-authentication-flow"></a><span data-ttu-id="39664-136">Поток проверки подлинности одного Sign-On</span><span class="sxs-lookup"><span data-stu-id="39664-136">Single Sign-On authentication flow</span></span>

> [!NOTE]
> <span data-ttu-id="39664-137">В настоящее время API единого входа поддерживается для Word, Excel, Outlook и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="39664-137">The single sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="39664-138">Дополнительные сведения о поддержке единого входа можно найти в статье [наборы требований IdentityAPI](../reference/requirement-sets/identity-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="39664-138">For more information about single sign-on support, see [IdentityAPI requirement sets](../reference/requirement-sets/identity-api-requirement-sets.md).</span></span> <span data-ttu-id="39664-139">Если вы работаете с надстройкой Outlook, обязательно включите современную проверку подлинности для клиента Office 365.</span><span class="sxs-lookup"><span data-stu-id="39664-139">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="39664-140">Сведения о том, как это сделать, см. в статье [Exchange Online: как включить в клиенте современную проверку подлинности](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="39664-140">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="39664-141">Используйте единый вход для более плавной работы конечных пользователей.</span><span class="sxs-lookup"><span data-stu-id="39664-141">Use single sign-on for a smoother end-user experience.</span></span> <span data-ttu-id="39664-142">Удостоверение пользователя в Office (учетная запись Майкрософт или удостоверение Microsoft 365) используется для входа в надстройку.</span><span class="sxs-lookup"><span data-stu-id="39664-142">The user's identity within Office (either a Microsoft Account or a Microsoft 365 identity) is used to sign in to your add-in.</span></span> <span data-ttu-id="39664-143">В результате пользователи только один раз войти.</span><span class="sxs-lookup"><span data-stu-id="39664-143">As a result users only sign in once.</span></span> <span data-ttu-id="39664-144">Это упрощает начало работы для пользователей.</span><span class="sxs-lookup"><span data-stu-id="39664-144">This removes friction in the experience making it easier for your customers to get started.</span></span>

1. <span data-ttu-id="39664-145">При установке надстройки пользователь увидит окно согласия, аналогичное следующему:</span><span class="sxs-lookup"><span data-stu-id="39664-145">As an add-in is being installed, a user will see a consent window similar to the one following:</span></span>

    ![Снимок экрана с окном согласия в приложении Office при установке надстройки](../images/add-in-auth-SSO-consent-dialog.png)

    > [!NOTE]
    > <span data-ttu-id="39664-147">Издатель надстройки может выбирать логотип, строки и разрешения, включаемые в окно запроса.</span><span class="sxs-lookup"><span data-stu-id="39664-147">The add-in publisher will have control over the logo, strings and permission scopes included in the consent window.</span></span> <span data-ttu-id="39664-148">Пользовательский интерфейс определяет Майкрософт.</span><span class="sxs-lookup"><span data-stu-id="39664-148">The UI is pre-configured by Microsoft.</span></span>

1. <span data-ttu-id="39664-149">Надстройка загрузится после того, как пользователь предоставит разрешения.</span><span class="sxs-lookup"><span data-stu-id="39664-149">The add-in will load after the user consents.</span></span> <span data-ttu-id="39664-150">Она может извлечь и отобразить необходимую персонализированную информацию.</span><span class="sxs-lookup"><span data-stu-id="39664-150">It can extract and display any necessary user customized information.</span></span>

    ![Снимок экрана с приложением Office с кнопками надстройки, отображаемыми на ленте](../images/add-in-ribbon.png)

## <a name="see-also"></a><span data-ttu-id="39664-152">См. также</span><span class="sxs-lookup"><span data-stu-id="39664-152">See also</span></span>

- <span data-ttu-id="39664-153">Дополнительные сведения о [разработке надстроек с единым входом](../develop/sso-in-office-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="39664-153">Learn more about [developing SSO Add-ins](../develop/sso-in-office-add-ins.md)</span></span>
