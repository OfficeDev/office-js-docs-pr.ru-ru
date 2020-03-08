---
title: Настройка надстройки с поддержкой единого входа для узла. js
description: Сведения о настройке надстройки с поддержкой единого входа, созданной с помощью генератора Yeoman.
ms.date: 02/20/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: c02e0f74a8ea3f3f8f831b65aa403ce49655953b
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/26/2020
ms.locfileid: "42284169"
---
# <a name="customize-your-nodejs-sso-enabled-add-in"></a><span data-ttu-id="bc84c-103">Настройка надстройки с поддержкой единого входа для узла. js</span><span class="sxs-lookup"><span data-stu-id="bc84c-103">Customize your Node.js SSO-enabled add-in</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bc84c-104">Эта статья основана на надстройке с поддержкой единого входа, которая создается с помощью краткого руководства по выполнению [единого входа (SSO)](sso-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="bc84c-104">This article builds upon the SSO-enabled add-in that's created by completing the [single sign-on (SSO) quick start](sso-quickstart.md).</span></span> <span data-ttu-id="bc84c-105">Прежде чем приступить к чтению этой статьи, заполните краткое руководство.</span><span class="sxs-lookup"><span data-stu-id="bc84c-105">Please complete the quick start before reading this article.</span></span>

<span data-ttu-id="bc84c-106">[Быстрый запуск единого входа](sso-quickstart.md) создает надстройку с включенной поддержкой единого входа, которая получает данные профиля пользователя, выполнившего вход, и записывает их в документ или сообщение.</span><span class="sxs-lookup"><span data-stu-id="bc84c-106">The [SSO quick start](sso-quickstart.md) creates an SSO-enabled add-in that gets the signed-in user's profile information and writes it to the document or message.</span></span> <span data-ttu-id="bc84c-107">В этой статье описывается процесс обновления надстройки, созданной с помощью генератора Yeoman в быстром запуске единого входа, для добавления новых функциональных возможностей, требующих других разрешений.</span><span class="sxs-lookup"><span data-stu-id="bc84c-107">In this article, you'll walk through the process of updating the add-in that you created with the Yeoman generator in the SSO quick start, to add new functionality that requires different permissions.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="bc84c-108">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="bc84c-108">Prerequisites</span></span>

* <span data-ttu-id="bc84c-109">Надстройка Office, созданная в соответствии с инструкциями, приведенными в [кратком](sso-quickstart.md)руководстве по SSO.</span><span class="sxs-lookup"><span data-stu-id="bc84c-109">An Office Add-in that you created by following the instructions in the [SSO quick start](sso-quickstart.md).</span></span>

* <span data-ttu-id="bc84c-110">Несколько файлов и папок, сохраненных в OneDrive для бизнеса в составе подписки на Office 365.</span><span class="sxs-lookup"><span data-stu-id="bc84c-110">At least a few files and folders stored on OneDrive for Business in your Office 365 subscription.</span></span>

* <span data-ttu-id="bc84c-111">[Node. js](https://nodejs.org) (последняя версия [LTS](https://nodejs.org/about/releases) ).</span><span class="sxs-lookup"><span data-stu-id="bc84c-111">[Node.js](https://nodejs.org) (the latest [LTS](https://nodejs.org/about/releases) version).</span></span>

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

## <a name="review-contents-of-the-project"></a><span data-ttu-id="bc84c-112">Просмотр содержимого проекта</span><span class="sxs-lookup"><span data-stu-id="bc84c-112">Review contents of the project</span></span>

<span data-ttu-id="bc84c-113">Начнем с краткого обзора проекта надстройки, [созданного ранее с помощью генератора Yeoman](sso-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="bc84c-113">Let's begin with a quick review of the add-in project that you previously [created with the Yeoman generator](sso-quickstart.md).</span></span>

> [!NOTE]
> <span data-ttu-id="bc84c-114">В местах, где эта статья ссылается на файлы сценариев с использованием расширения **JS** , вместо этого следует использовать расширение **TS** , если проект был создан с помощью TypeScript.</span><span class="sxs-lookup"><span data-stu-id="bc84c-114">In places where this article references script files using **.js** file extension, assume the **.ts** file extension instead if your project was created with TypeScript.</span></span>

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## <a name="add-new-functionality"></a><span data-ttu-id="bc84c-115">Добавление новых функциональных возможностей</span><span class="sxs-lookup"><span data-stu-id="bc84c-115">Add new functionality</span></span> 

<span data-ttu-id="bc84c-116">Надстройка, созданная с помощью быстрого запуска единого входа, использует Microsoft Graph для получения сведений о профиле пользователя, выполнившего вход, и записывает эти сведения в документ или сообщение.</span><span class="sxs-lookup"><span data-stu-id="bc84c-116">The add-in that you created with the SSO quick start uses Microsoft Graph to get the signed-in user's profile information and writes that information to the document or message.</span></span> <span data-ttu-id="bc84c-117">Теперь изменим функциональные возможности надстройки, чтобы она выводила имена 10 файлов и папок из OneDrive для бизнеса пользователя, выполнившего вход, и записывает эти сведения в документ или сообщение.</span><span class="sxs-lookup"><span data-stu-id="bc84c-117">Let's change the add-in's functionality such that it gets the names of the top 10 files and folders from the signed-in user's OneDrive for Business and writes that information to the document or message.</span></span> <span data-ttu-id="bc84c-118">Для этого требуется обновление разрешений приложений в Azure и обновление кода в проекте надстройки.</span><span class="sxs-lookup"><span data-stu-id="bc84c-118">Enabling this new functionality requires updating app permissions in Azure and updating code within the add-in project.</span></span>

### <a name="update-app-permissions-in-azure"></a><span data-ttu-id="bc84c-119">Обновление разрешений приложения в Azure</span><span class="sxs-lookup"><span data-stu-id="bc84c-119">Update app permissions in Azure</span></span>

<span data-ttu-id="bc84c-120">Прежде чем надстройка сможет успешно прочитать содержимое OneDrive для бизнеса пользователя, ее регистрационная информация в Azure должна быть обновлена с соответствующими разрешениями.</span><span class="sxs-lookup"><span data-stu-id="bc84c-120">Before the add-in can successfully read the contents of the user's OneDrive for Business, its app registration information in Azure must be updated with the appropriate permissions.</span></span> <span data-ttu-id="bc84c-121">Выполните следующие действия, чтобы предоставить приложению разрешение **Files. Read. ALL** и отозвать разрешение **User.** Read. ALL, что больше не требуется.</span><span class="sxs-lookup"><span data-stu-id="bc84c-121">Complete the following steps to grant the app the **Files.Read.All** permission and revoke the **User.Read** permission, which is no longer needed.</span></span>

1. <span data-ttu-id="bc84c-122">Перейдите на [портал Azure](https://ms.portal.azure.com/#home) и **Войдите в систему, используя учетные данные администратора Office 365**.</span><span class="sxs-lookup"><span data-stu-id="bc84c-122">Navigate to the [Azure portal](https://ms.portal.azure.com/#home) and **sign in using your Office 365 administrator credentials**.</span></span> 

2. <span data-ttu-id="bc84c-123">Перейдите на страницу **регистрации приложений** .</span><span class="sxs-lookup"><span data-stu-id="bc84c-123">Navigate to the **App registrations** page.</span></span> 
    > [!TIP]
    > <span data-ttu-id="bc84c-124">Это можно сделать, выбрав плитку **регистрации приложений** на домашней странице Azure или воспользовавшись полем поиска на домашней странице, чтобы найти и выбрать **регистрации приложений**.</span><span class="sxs-lookup"><span data-stu-id="bc84c-124">You can do this either by choosing the **App registrations** tile on the Azure home page or by using the search box on the home page to find and choose **App registrations**.</span></span>

3. <span data-ttu-id="bc84c-125">На странице **регистрации приложений** выберите приложение, созданное на этапе быстрого запуска.</span><span class="sxs-lookup"><span data-stu-id="bc84c-125">On the **App registrations** page, choose the app that you created during the quick start.</span></span> 
    > [!TIP]
    > <span data-ttu-id="bc84c-126">**Отображаемое имя** приложения будет соответствующим имени надстройки, которое вы указали при создании проекта с помощью генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="bc84c-126">The **Display name** of the app will match the add-in name that you specified when you created the project with the Yeoman generator.</span></span>

4. <span data-ttu-id="bc84c-127">На странице "Обзор приложения" выберите **разрешения API** в разделе **Управление** заголовком в левой части страницы.</span><span class="sxs-lookup"><span data-stu-id="bc84c-127">From the app overview page, choose **API permissions** under the **Manage** heading on the left side of the page.</span></span>

5. <span data-ttu-id="bc84c-128">В строке **User. Read** таблицы Permissions нажмите кнопку с многоточием, а затем выберите **отозвать согласие администратора** из появившегося меню.</span><span class="sxs-lookup"><span data-stu-id="bc84c-128">In the **User.Read** row of the permissions table, choose the ellipsis and then select **Revoke admin consent** from the menu that appears.</span></span>

6. <span data-ttu-id="bc84c-129">Нажмите кнопку **Да, удалить** в ответ на отображаемый запрос.</span><span class="sxs-lookup"><span data-stu-id="bc84c-129">Select the **Yes, remove** button in response to the prompt that's displayed.</span></span>

7. <span data-ttu-id="bc84c-130">В строке **User. Read** таблицы Permissions нажмите кнопку с многоточием, а затем выберите пункт **удалить разрешение** из появившегося меню.</span><span class="sxs-lookup"><span data-stu-id="bc84c-130">In the **User.Read** row of the permissions table, choose the ellipsis and then select **Remove permission** from the menu that appears.</span></span>

8. <span data-ttu-id="bc84c-131">Нажмите кнопку **Да, удалить** в ответ на отображаемый запрос.</span><span class="sxs-lookup"><span data-stu-id="bc84c-131">Select the **Yes, remove** button in response to the prompt that's displayed.</span></span>

9. <span data-ttu-id="bc84c-132">Нажмите кнопку **Добавить разрешение** .</span><span class="sxs-lookup"><span data-stu-id="bc84c-132">Select the **Add a permission** button.</span></span>

10. <span data-ttu-id="bc84c-133">В открывшейся панели выберите **Microsoft Graph** , а затем — **делегированные разрешения**.</span><span class="sxs-lookup"><span data-stu-id="bc84c-133">On the panel that opens choose **Microsoft Graph** and then choose **Delegated permissions**.</span></span>

11. <span data-ttu-id="bc84c-134">На панели **разрешений API запроса** выполните следующие действия:</span><span class="sxs-lookup"><span data-stu-id="bc84c-134">On the **Request API permissions** panel:</span></span>

    <span data-ttu-id="bc84c-135">а.</span><span class="sxs-lookup"><span data-stu-id="bc84c-135">a.</span></span> <span data-ttu-id="bc84c-136">В разделе **файлы**выберите **файлы. Read. ALL**.</span><span class="sxs-lookup"><span data-stu-id="bc84c-136">Under **Files**, select **Files.Read.All**.</span></span>

    <span data-ttu-id="bc84c-137">б.</span><span class="sxs-lookup"><span data-stu-id="bc84c-137">b.</span></span> <span data-ttu-id="bc84c-138">Нажмите кнопку **Добавить разрешения** в нижней части панели, чтобы сохранить изменения этих разрешений.</span><span class="sxs-lookup"><span data-stu-id="bc84c-138">Select the **Add permissions** button at the bottom of the panel to save these permissions changes.</span></span>

12. <span data-ttu-id="bc84c-139">Нажмите кнопку **предоставить согласие администратора для пользователя [имя клиента]** .</span><span class="sxs-lookup"><span data-stu-id="bc84c-139">Select the **Grant admin consent for [tenant name]** button.</span></span>

13. <span data-ttu-id="bc84c-140">Нажмите кнопку **Да** в ответ на отображаемый запрос.</span><span class="sxs-lookup"><span data-stu-id="bc84c-140">Select the **Yes** button in response to the prompt that's displayed.</span></span>

### <a name="update-code-in-the-add-in-project"></a><span data-ttu-id="bc84c-141">Обновление кода в проекте надстройки</span><span class="sxs-lookup"><span data-stu-id="bc84c-141">Update code in the add-in project</span></span>

<span data-ttu-id="bc84c-142">Чтобы надстройка прочитала содержимое OneDrive для бизнеса пользователя, выполнившего вход, необходимо выполнить следующие действия:</span><span class="sxs-lookup"><span data-stu-id="bc84c-142">To enable the add-in to read contents of the signed-in user's OneDrive for Business, you'll need to:</span></span>

- <span data-ttu-id="bc84c-143">Обновите код, ссылающийся на URL-адрес, параметры и требуемую область доступа Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="bc84c-143">Update the code that references the Microsoft Graph URL, parameters, and required access scope.</span></span>

- <span data-ttu-id="bc84c-144">Обновите код, определяющий пользовательский интерфейс области задач, чтобы он точно описывает новые функциональные возможности.</span><span class="sxs-lookup"><span data-stu-id="bc84c-144">Update the code that defines the task pane UI, so that it accurately describes the new functionality.</span></span> 

- <span data-ttu-id="bc84c-145">Обновление кода, который анализирует отклик от Microsoft Graph и записывает его в документ или сообщение.</span><span class="sxs-lookup"><span data-stu-id="bc84c-145">Update the code that parses the response from Microsoft Graph and writes it to the document or message.</span></span>

<span data-ttu-id="bc84c-146">Эти обновления описываются в следующих шагах.</span><span class="sxs-lookup"><span data-stu-id="bc84c-146">The following steps describe these updates.</span></span>

### <a name="changes-required-for-any-type-of-add-in"></a><span data-ttu-id="bc84c-147">Изменения, необходимые для любого типа надстройки</span><span class="sxs-lookup"><span data-stu-id="bc84c-147">Changes required for any type of add-in</span></span>

<span data-ttu-id="bc84c-148">Выполните указанные ниже действия для надстройки, чтобы изменить URL-адрес, параметры и область доступа Microsoft Graph, а также обновить пользовательский интерфейс области задач.</span><span class="sxs-lookup"><span data-stu-id="bc84c-148">Complete the following steps for your add-in, to change the Microsoft Graph URL, parameters, and access scope, and update the taskpane UI.</span></span> <span data-ttu-id="bc84c-149">Эти действия одинаковы, независимо от того, в каком приложении Office размещены целевые объекты надстройки.</span><span class="sxs-lookup"><span data-stu-id="bc84c-149">These steps are the same, regardless of which Office host your add-in targets.</span></span>

1. <span data-ttu-id="bc84c-150">В файле **./. ENV** :</span><span class="sxs-lookup"><span data-stu-id="bc84c-150">In the **./.ENV** file:</span></span>

    <span data-ttu-id="bc84c-151">а.</span><span class="sxs-lookup"><span data-stu-id="bc84c-151">a.</span></span> <span data-ttu-id="bc84c-152">Замените `GRAPH_URL_SEGMENT=/me` на следующий:`GRAPH_URL_SEGMENT=/me/drive/root/children`</span><span class="sxs-lookup"><span data-stu-id="bc84c-152">Replace `GRAPH_URL_SEGMENT=/me` with the following: `GRAPH_URL_SEGMENT=/me/drive/root/children`</span></span>

    <span data-ttu-id="bc84c-153">б.</span><span class="sxs-lookup"><span data-stu-id="bc84c-153">b.</span></span> <span data-ttu-id="bc84c-154">Замените `QUERY_PARAM_SEGMENT=` на следующий:`QUERY_PARAM_SEGMENT=?$select=name&$top=10`</span><span class="sxs-lookup"><span data-stu-id="bc84c-154">Replace `QUERY_PARAM_SEGMENT=` with the following: `QUERY_PARAM_SEGMENT=?$select=name&$top=10`</span></span>

    <span data-ttu-id="bc84c-155">в.</span><span class="sxs-lookup"><span data-stu-id="bc84c-155">c.</span></span> <span data-ttu-id="bc84c-156">Замените `SCOPE=User.Read` на следующий:`SCOPE=Files.Read.All`</span><span class="sxs-lookup"><span data-stu-id="bc84c-156">Replace `SCOPE=User.Read` with the following: `SCOPE=Files.Read.All`</span></span>

2. <span data-ttu-id="bc84c-157">В **/манифест.ксмл**найдите строку `<Scope>User.Read</Scope>` около конца файла и замените ее на строку `<Scope>Files.Read.All</Scope>`.</span><span class="sxs-lookup"><span data-stu-id="bc84c-157">In **./manifest.xml**, find the line `<Scope>User.Read</Scope>` near the end of the file and replace it with the line `<Scope>Files.Read.All</Scope>`.</span></span>

3. <span data-ttu-id="bc84c-158">В файле **./СРК/Хелперс/фаллбаккаусдиалог.ЖС** (или **в/СРК/Хелперс/фаллбаккаусдиалог.ТС** для проекта TypeScript) найдите `https://graph.microsoft.com/User.Read` строку и замените ее строкой `https://graph.microsoft.com/Files.Read.All`, которая `requestObj` определяется следующим образом:</span><span class="sxs-lookup"><span data-stu-id="bc84c-158">In **./src/helpers/fallbackauthdialog.js** (or in **./src/helpers/fallbackauthdialog.ts** for a TypeScript project), find the string `https://graph.microsoft.com/User.Read` and replace it with the string `https://graph.microsoft.com/Files.Read.All`, such that `requestObj` is defined as follows:</span></span>

    ```javascript
    var requestObj = {
      scopes: [`https://graph.microsoft.com/Files.Read.All`]
    };
    ```

    ```typescript
    var requestObj: Object = {
      scopes: [`https://graph.microsoft.com/Files.Read.All`]
    };
    ```

4. <span data-ttu-id="bc84c-159">В **src/TaskPane/TaskPane.HTML**найдите элемент `<section class="ms-firstrun-instructionstep__header">` и обновите текст в этом элементе, чтобы описать новые функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="bc84c-159">In **./src/taskpane/taskpane.html**, find the element `<section class="ms-firstrun-instructionstep__header">` and update the text within that element to describe the add-in's new functionality.</span></span>

    ```html
    <section class="ms-firstrun-instructionstep__header">
        <h2 class="ms-font-m">This add-in demonstrates how to use single sign-on by making a call to Microsoft
            Graph to read content from OneDrive for Business.</h2>
        <div class="ms-firstrun-instructionstep__header--image"></div>
    </section>
    ```

5. <span data-ttu-id="bc84c-160">В файле **./src/TaskPane/TaskPane.HTML**найдите и замените все вхождения строки `Get My User Profile Information` строкой. `Read my OneDrive for Business`</span><span class="sxs-lookup"><span data-stu-id="bc84c-160">In **./src/taskpane/taskpane.html**, find and replace both occurrences of the string `Get My User Profile Information` with the string `Read my OneDrive for Business`.</span></span>

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">Click the <b>Read my OneDrive for Business</b>
            button.</span>
        <div class="clearfix"></div>
    </li>
    ```

    ```html
    <p align="center">
        <button id="getGraphDataButton" class="popupButton ms-Button ms-Button--primary"><span
                class="ms-Button-label">Read my OneDrive for Business</span></button>
    </p>
    ```

6. <span data-ttu-id="bc84c-161">В файле **./src/TaskPane/TaskPane.HTML**найдите и замените строку `Your user profile information will be displayed in the document.` строкой. `The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.`</span><span class="sxs-lookup"><span data-stu-id="bc84c-161">In **./src/taskpane/taskpane.html**, find and replace the string `Your user profile information will be displayed in the document.` with the string `The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.`.</span></span>

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.</span>
        <div class="clearfix"></div>
    </li>
    ```

7. <span data-ttu-id="bc84c-162">Обновите код, который анализирует ответ от Microsoft Graph, и записывает его в документ или сообщение, следуя указаниям в разделе, соответствующем типу надстройки:</span><span class="sxs-lookup"><span data-stu-id="bc84c-162">Update the code that parses the response from Microsoft Graph and writes it to the document or message, by following guidance in the section that corresponds to your type of add-in:</span></span>

    - [<span data-ttu-id="bc84c-163">Изменения, необходимые для надстройки Excel (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="bc84c-163">Changes required for an Excel add-in (JavaScript)</span></span>](#changes-required-for-an-excel-add-in-javascript)
    - [<span data-ttu-id="bc84c-164">Изменения, необходимые для надстройки Excel (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="bc84c-164">Changes required for an Excel add-in (TypeScript)</span></span>](#changes-required-for-an-excel-add-in-typescript)
    - [<span data-ttu-id="bc84c-165">Изменения, необходимые для надстройки Outlook (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="bc84c-165">Changes required for an Outlook add-in (JavaScript)</span></span>](#changes-required-for-an-outlook-add-in-javascript)
    - [<span data-ttu-id="bc84c-166">Изменения, необходимые для надстройки Outlook (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="bc84c-166">Changes required for an Outlook add-in (TypeScript)</span></span>](#changes-required-for-an-outlook-add-in-typescript)
    - [<span data-ttu-id="bc84c-167">Изменения, необходимые для надстройки PowerPoint (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="bc84c-167">Changes required for a PowerPoint add-in (JavaScript)</span></span>](#changes-required-for-a-powerpoint-add-in-javascript)
    - [<span data-ttu-id="bc84c-168">Изменения, необходимые для надстройки PowerPoint (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="bc84c-168">Changes required for a PowerPoint add-in (TypeScript)</span></span>](#changes-required-for-a-powerpoint-add-in-typescript)
    - [<span data-ttu-id="bc84c-169">Изменения, необходимые для надстройки Word (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="bc84c-169">Changes required for a Word add-in (JavaScript)</span></span>](#changes-required-for-a-word-add-in-javascript)
    - [<span data-ttu-id="bc84c-170">Изменения, необходимые для надстройки Word (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="bc84c-170">Changes required for a Word add-in (TypeScript)</span></span>](#changes-required-for-a-word-add-in-typescript)

### <a name="changes-required-for-an-excel-add-in-javascript"></a><span data-ttu-id="bc84c-171">Изменения, необходимые для надстройки Excel (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="bc84c-171">Changes required for an Excel add-in (JavaScript)</span></span>

<span data-ttu-id="bc84c-172">Если надстройка представляет собой надстройку Excel, созданную с помощью JavaScript, внесите следующие изменения в файле **./СРК/Хелперс/докуменселпер.ЖС**:</span><span class="sxs-lookup"><span data-stu-id="bc84c-172">If your add-in is an Excel add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="bc84c-173">Найдите `writeDataToOfficeDocument` функцию и замените ее следующей функцией:</span><span class="sxs-lookup"><span data-stu-id="bc84c-173">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToExcel(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. <span data-ttu-id="bc84c-174">Найдите `filterUserProfileInfo` функцию и замените ее следующей функцией:</span><span class="sxs-lookup"><span data-stu-id="bc84c-174">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. <span data-ttu-id="bc84c-175">Найдите `writeDataToExcel` функцию и замените ее следующей функцией:</span><span class="sxs-lookup"><span data-stu-id="bc84c-175">Find the `writeDataToExcel` function and replace it with the following function:</span></span>

    ```javascript
    function writeDataToExcel(result) {
      return Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        let data = [];
        let oneDriveInfo = filterOneDriveInfo(result);

        for (let i = 0; i < oneDriveInfo.length; i++) {
          if (oneDriveInfo[i] !== null) {
            let innerArray = [];
            innerArray.push(oneDriveInfo[i]);
            data.push(innerArray);
          }
        }

        const rangeAddress = `B5:B${5 + (data.length - 1)}`;
        const range = sheet.getRange(rangeAddress);
        range.values = data;
        range.format.autofitColumns();

        return context.sync();
      });
    }
    ```

4. <span data-ttu-id="bc84c-176">Удалите `writeDataToOutlook` функцию.</span><span class="sxs-lookup"><span data-stu-id="bc84c-176">Delete the `writeDataToOutlook` function.</span></span>

5. <span data-ttu-id="bc84c-177">Удалите `writeDataToPowerPoint` функцию.</span><span class="sxs-lookup"><span data-stu-id="bc84c-177">Delete the `writeDataToPowerPoint` function.</span></span>

6. <span data-ttu-id="bc84c-178">Удалите `writeDataToWord` функцию.</span><span class="sxs-lookup"><span data-stu-id="bc84c-178">Delete the `writeDataToWord` function.</span></span>

<span data-ttu-id="bc84c-179">После внесения этих изменений перейдите к разделу " [попробовать](#try-it-out) " в этой статье, чтобы испытать обновленную надстройку.</span><span class="sxs-lookup"><span data-stu-id="bc84c-179">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-an-excel-add-in-typescript"></a><span data-ttu-id="bc84c-180">Изменения, необходимые для надстройки Excel (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="bc84c-180">Changes required for an Excel add-in (TypeScript)</span></span>

<span data-ttu-id="bc84c-181">Если надстройка представляет собой надстройку Excel, созданную с помощью TypeScript, откройте **./СРК/таскпане/таскпане.ТС**, найдите `writeDataToOfficeDocument` функцию и замените ее следующей функцией:</span><span class="sxs-lookup"><span data-stu-id="bc84c-181">If your add-in is an Excel add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

```typescript
export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Excel.run(function(context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
      itemNames.push(item["name"]);
    }

    for (let i = 0; i < itemNames.length; i++) {
      if (itemNames[i] !== null) {
        let innerArray = [];
        innerArray.push(itemNames[i]);
        data.push(innerArray);
      }
    }
    
    const rangeAddress = `B5:B${5 + (data.length - 1)}`;
    const range = sheet.getRange(rangeAddress);
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
  });
}
```

<span data-ttu-id="bc84c-182">После внесения этих изменений перейдите к разделу " [попробовать](#try-it-out) " в этой статье, чтобы испытать обновленную надстройку.</span><span class="sxs-lookup"><span data-stu-id="bc84c-182">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-an-outlook-add-in-javascript"></a><span data-ttu-id="bc84c-183">Изменения, необходимые для надстройки Outlook (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="bc84c-183">Changes required for an Outlook add-in (JavaScript)</span></span>

<span data-ttu-id="bc84c-184">Если надстройка представляет собой надстройку Outlook, созданную с помощью JavaScript, внесите следующие изменения в файле **./СРК/Хелперс/докуменселпер.ЖС**:</span><span class="sxs-lookup"><span data-stu-id="bc84c-184">If your add-in is an Outlook add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="bc84c-185">Найдите `writeDataToOfficeDocument` функцию и замените ее следующей функцией:</span><span class="sxs-lookup"><span data-stu-id="bc84c-185">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToOutlook(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to message. " + error.toString()));
        }
      });
    }
    ```

2. <span data-ttu-id="bc84c-186">Найдите `filterUserProfileInfo` функцию и замените ее следующей функцией:</span><span class="sxs-lookup"><span data-stu-id="bc84c-186">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. <span data-ttu-id="bc84c-187">Найдите `writeDataToOutlook` функцию и замените ее следующей функцией:</span><span class="sxs-lookup"><span data-stu-id="bc84c-187">Find the `writeDataToOutlook` function and replace it with the following function:</span></span>

    ```javascript
    function writeDataToOutlook(result) {
      let data = [];
      let oneDriveInfo = filterOneDriveInfo(result);

      for (let i = 0; i < oneDriveInfo.length; i++) {
        if (oneDriveInfo[i] !== null) {
          data.push(oneDriveInfo[i]);
        }
      }

      let objectNames = "";
      for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "<br/>";
      }

      Office.context.mailbox.item.body.setSelectedDataAsync(objectNames, { coercionType: Office.CoercionType.Html });
    }
    ```

4. <span data-ttu-id="bc84c-188">Удалите `writeDataToExcel` функцию.</span><span class="sxs-lookup"><span data-stu-id="bc84c-188">Delete the `writeDataToExcel` function.</span></span>

5. <span data-ttu-id="bc84c-189">Удалите `writeDataToPowerPoint` функцию.</span><span class="sxs-lookup"><span data-stu-id="bc84c-189">Delete the `writeDataToPowerPoint` function.</span></span>

6. <span data-ttu-id="bc84c-190">Удалите `writeDataToWord` функцию.</span><span class="sxs-lookup"><span data-stu-id="bc84c-190">Delete the `writeDataToWord` function.</span></span>

<span data-ttu-id="bc84c-191">После внесения этих изменений перейдите к разделу " [попробовать](#try-it-out) " в этой статье, чтобы испытать обновленную надстройку.</span><span class="sxs-lookup"><span data-stu-id="bc84c-191">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-an-outlook-add-in-typescript"></a><span data-ttu-id="bc84c-192">Изменения, необходимые для надстройки Outlook (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="bc84c-192">Changes required for an Outlook add-in (TypeScript)</span></span>

<span data-ttu-id="bc84c-193">Если надстройка представляет собой надстройку Outlook, созданную с помощью TypeScript, откройте **./СРК/таскпане/таскпане.ТС**, найдите `writeDataToOfficeDocument` функцию и замените ее следующей функцией:</span><span class="sxs-lookup"><span data-stu-id="bc84c-193">If your add-in is an Outlook add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

```typescript
export function writeDataToOfficeDocument(result: Object): void {
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
        itemNames.push(item["name"]);
    };

    for (let i = 0; i < itemNames.length; i++) {
        if (itemNames[i] !== null) {
        data.push(itemNames[i]);
        }
    }

    let objectNames: string = "";
    for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "<br/>";
    }
    
    Office.context.mailbox.item.body.setSelectedDataAsync(objectNames, { coercionType: Office.CoercionType.Html });
}
```

<span data-ttu-id="bc84c-194">После внесения этих изменений перейдите к разделу " [попробовать](#try-it-out) " в этой статье, чтобы испытать обновленную надстройку.</span><span class="sxs-lookup"><span data-stu-id="bc84c-194">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-powerpoint-add-in-javascript"></a><span data-ttu-id="bc84c-195">Изменения, необходимые для надстройки PowerPoint (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="bc84c-195">Changes required for a PowerPoint add-in (JavaScript)</span></span>

<span data-ttu-id="bc84c-196">Если надстройка представляет собой надстройку PowerPoint, созданную с помощью JavaScript, внесите следующие изменения в файле **./СРК/Хелперс/докуменселпер.ЖС**:</span><span class="sxs-lookup"><span data-stu-id="bc84c-196">If your add-in is a PowerPoint add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="bc84c-197">Найдите `writeDataToOfficeDocument` функцию и замените ее следующей функцией:</span><span class="sxs-lookup"><span data-stu-id="bc84c-197">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToPowerPoint(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. <span data-ttu-id="bc84c-198">Найдите `filterUserProfileInfo` функцию и замените ее следующей функцией:</span><span class="sxs-lookup"><span data-stu-id="bc84c-198">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. <span data-ttu-id="bc84c-199">Найдите `writeDataToPowerPoint` функцию и замените ее следующей функцией:</span><span class="sxs-lookup"><span data-stu-id="bc84c-199">Find the `writeDataToPowerPoint` function and replace it with the following function:</span></span>

    ```javascript
    function writeDataToPowerPoint(result) {
      let data = [];
      let oneDriveInfo = filterOneDriveInfo(result);

      for (let i = 0; i < oneDriveInfo.length; i++) {
        if (oneDriveInfo[i] !== null) {
          data.push(oneDriveInfo[i]);
        }
      }

      let objectNames = "";
      for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "\n";
      }

      Office.context.document.setSelectedDataAsync(
        objectNames, 
        function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            throw asyncResult.error.message;
          }
      });
    }
    ```

4. <span data-ttu-id="bc84c-200">Удалите `writeDataToExcel` функцию.</span><span class="sxs-lookup"><span data-stu-id="bc84c-200">Delete the `writeDataToExcel` function.</span></span>

5. <span data-ttu-id="bc84c-201">Удалите `writeDataToOutlook` функцию.</span><span class="sxs-lookup"><span data-stu-id="bc84c-201">Delete the `writeDataToOutlook` function.</span></span>

6. <span data-ttu-id="bc84c-202">Удалите `writeDataToWord` функцию.</span><span class="sxs-lookup"><span data-stu-id="bc84c-202">Delete the `writeDataToWord` function.</span></span>

<span data-ttu-id="bc84c-203">После внесения этих изменений перейдите к разделу " [попробовать](#try-it-out) " в этой статье, чтобы испытать обновленную надстройку.</span><span class="sxs-lookup"><span data-stu-id="bc84c-203">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-powerpoint-add-in-typescript"></a><span data-ttu-id="bc84c-204">Изменения, необходимые для надстройки PowerPoint (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="bc84c-204">Changes required for a PowerPoint add-in (TypeScript)</span></span>

<span data-ttu-id="bc84c-205">Если надстройка представляет собой надстройку PowerPoint, созданную с помощью TypeScript, откройте **./СРК/таскпане/таскпане.ТС**, найдите `writeDataToOfficeDocument` функцию и замените ее следующей функцией:</span><span class="sxs-lookup"><span data-stu-id="bc84c-205">If your add-in is a PowerPoint add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

```typescript
export function writeDataToOfficeDocument(result: Object): void {
  let data: string[] = [];

  let itemNames: string[] = [];
  let oneDriveItems = result["value"];
  for (let item of oneDriveItems) {
    itemNames.push(item["name"]);
  };

  for (let i = 0; i < itemNames.length; i++) {
    if (itemNames[i] !== null) {
      data.push(itemNames[i]);
    }
  }

  let objectNames: string = "";
  for (let i = 0; i < data.length; i++) {
    objectNames += data[i] + "\n";
  }

  Office.context.document.setSelectedDataAsync(objectNames, function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      throw asyncResult.error.message;
    }
  });
}
```

<span data-ttu-id="bc84c-206">После внесения этих изменений перейдите к разделу " [попробовать](#try-it-out) " в этой статье, чтобы испытать обновленную надстройку.</span><span class="sxs-lookup"><span data-stu-id="bc84c-206">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-word-add-in-javascript"></a><span data-ttu-id="bc84c-207">Изменения, необходимые для надстройки Word (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="bc84c-207">Changes required for a Word add-in (JavaScript)</span></span>

<span data-ttu-id="bc84c-208">Если надстройка представляет собой надстройку Word, созданную с помощью JavaScript, внесите следующие изменения в файле **./СРК/Хелперс/докуменселпер.ЖС**:</span><span class="sxs-lookup"><span data-stu-id="bc84c-208">If your add-in is a Word add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="bc84c-209">Найдите `writeDataToOfficeDocument` функцию и замените ее следующей функцией:</span><span class="sxs-lookup"><span data-stu-id="bc84c-209">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToWord(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. <span data-ttu-id="bc84c-210">Найдите `filterUserProfileInfo` функцию и замените ее следующей функцией:</span><span class="sxs-lookup"><span data-stu-id="bc84c-210">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. <span data-ttu-id="bc84c-211">Найдите `writeDataToWord` функцию и замените ее следующей функцией:</span><span class="sxs-lookup"><span data-stu-id="bc84c-211">Find the `writeDataToWord` function and replace it with the following function:</span></span>

    ```javascript
    function writeDataToWord(result) {
      return Word.run(function (context) {
        let data = [];
        let oneDriveInfo = filterOneDriveInfo(result);

        for (let i = 0; i < oneDriveInfo.length; i++) {
          if (oneDriveInfo[i] !== null) {
            data.push(oneDriveInfo[i]);
          }
        }

        const documentBody = context.document.body;
        for (let i = 0; i < data.length; i++) {
          if (data[i] !== null) {
            documentBody.insertParagraph(data[i], "End");
          }
        }

        return context.sync();
      });
    }
    ```

4. <span data-ttu-id="bc84c-212">Удалите `writeDataToExcel` функцию.</span><span class="sxs-lookup"><span data-stu-id="bc84c-212">Delete the `writeDataToExcel` function.</span></span>

5. <span data-ttu-id="bc84c-213">Удалите `writeDataToOutlook` функцию.</span><span class="sxs-lookup"><span data-stu-id="bc84c-213">Delete the `writeDataToOutlook` function.</span></span>

6. <span data-ttu-id="bc84c-214">Удалите `writeDataToPowerPoint` функцию.</span><span class="sxs-lookup"><span data-stu-id="bc84c-214">Delete the `writeDataToPowerPoint` function.</span></span>

<span data-ttu-id="bc84c-215">После внесения этих изменений перейдите к разделу " [попробовать](#try-it-out) " в этой статье, чтобы испытать обновленную надстройку.</span><span class="sxs-lookup"><span data-stu-id="bc84c-215">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-word-add-in-typescript"></a><span data-ttu-id="bc84c-216">Изменения, необходимые для надстройки Word (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="bc84c-216">Changes required for a Word add-in (TypeScript)</span></span>

<span data-ttu-id="bc84c-217">Если надстройка представляет собой надстройку Word, созданную с помощью TypeScript, откройте **./СРК/таскпане/таскпане.ТС**, найдите `writeDataToOfficeDocument` функцию и замените ее следующей функцией:</span><span class="sxs-lookup"><span data-stu-id="bc84c-217">If your add-in is a Word add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

```typescript
export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Word.run(function(context) {
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
      itemNames.push(item["name"]);
    };

    for (let i = 0; i < itemNames.length; i++) {
      if (itemNames[i] !== null) {
        data.push(itemNames[i]);
      }
    }

    const documentBody: Word.Body = context.document.body;
    for (let i = 0; i < data.length; i++) {
      if (data[i] !== null) {
        documentBody.insertParagraph(data[i], "End");
      }
    }
    return context.sync();
  });
}
```

<span data-ttu-id="bc84c-218">После внесения этих изменений перейдите [к разделу](#try-it-out) "ознакомьтесь с этой статьей", чтобы испытать обновленную надстройку.</span><span class="sxs-lookup"><span data-stu-id="bc84c-218">After you've made these changes, continue to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="bc84c-219">Проверка</span><span class="sxs-lookup"><span data-stu-id="bc84c-219">Try it out</span></span>

<span data-ttu-id="bc84c-220">Если надстройка представляет собой надстройку Excel, Word или PowerPoint, выполните действия, описанные в следующем разделе, чтобы попробовать. Если надстройка является надстройкой Outlook, выполните действия, описанные в разделе [Outlook](#outlook) .</span><span class="sxs-lookup"><span data-stu-id="bc84c-220">If your add-in is an Excel, Word, or PowerPoint add-in, complete the steps in the following section to try it out. If your add-in is an Outlook add-in, complete the steps in the [Outlook](#outlook) section instead.</span></span>

### <a name="excel-word-and-powerpoint"></a><span data-ttu-id="bc84c-221">Excel, Word и PowerPoint</span><span class="sxs-lookup"><span data-stu-id="bc84c-221">Excel, Word, and PowerPoint</span></span>

<span data-ttu-id="bc84c-222">Выполните следующие действия, чтобы испытать надстройку Excel, Word или PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="bc84c-222">Complete the following steps to try out an Excel, Word, or PowerPoint add-in.</span></span>

1. <span data-ttu-id="bc84c-223">В корневой папке проекта выполните следующую команду, чтобы выполнить сборку проекта, запустите локальный веб-сервер и Загрузка неопубликованных вашу надстройку в выбранном ранее клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="bc84c-223">In the root folder of the project, run the following command to build the project, start the local web server, and sideload your add-in in the previously selected Office client application.</span></span>

    > [!NOTE]
    > <span data-ttu-id="bc84c-224">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="bc84c-224">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="bc84c-225">Если вам будет предложено установить сертификат после того, как вы запустите указанную ниже команду, примите предложение установить сертификат, предоставленный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="bc84c-225">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="bc84c-226">В клиентском приложении Office, которое открывается при выполнении предыдущей команды (например, Excel, Word или PowerPoint), убедитесь, что вы вошли в систему с учетной записью пользователя, который является участником той же организации Office 365, что и учетная запись администратора Office 365, которую вы использовали для подключения к Azure при [настройке единого входа](sso-quickstart.md#configure-sso) для приложения.</span><span class="sxs-lookup"><span data-stu-id="bc84c-226">In the Office client application that opens when you run the previous command (i.e., Excel, Word or PowerPoint), make sure that you're signed in with a user that's a member of the same Office 365 organization as the Office 365 administrator account that you used to connect to Azure while [configuring SSO](sso-quickstart.md#configure-sso) for the app.</span></span> <span data-ttu-id="bc84c-227">Благодаря этому будут созданы соответствующие условия для успешного единого входа.</span><span class="sxs-lookup"><span data-stu-id="bc84c-227">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="bc84c-228">В клиентском приложении Office выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="bc84c-228">In the Office client application, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="bc84c-229">На рисунке ниже показана эта кнопка в Excel. </span><span class="sxs-lookup"><span data-stu-id="bc84c-229">The following image shows this button in Excel.</span></span>

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="bc84c-231">В нижней части области задач нажмите кнопку **прочитать мою службу OneDrive для бизнеса** , чтобы начать процесс единого входа.</span><span class="sxs-lookup"><span data-stu-id="bc84c-231">At the bottom of the task pane, choose the **Read my OneDrive for Business** button to initiate the SSO process.</span></span> 

5. <span data-ttu-id="bc84c-232">Если открывается диалоговое окно, в котором запрашиваются разрешения от имени надстройки, это означает, что единый вход не поддерживается для вашего сценария и надстройка использует альтернативный метод проверки подлинности пользователя.</span><span class="sxs-lookup"><span data-stu-id="bc84c-232">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="bc84c-233">Это может произойти, если администратор клиента не дал согласие на доступ надстройки к Microsoft Graph или если пользователь не вошел в Office с помощью действительной учетной записи Майкрософт или Office 365 (рабочей или учебной учетной записи).</span><span class="sxs-lookup"><span data-stu-id="bc84c-233">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft Account or Office 365 ("Work or School") account.</span></span> <span data-ttu-id="bc84c-234">Чтобы продолжить, нажмите кнопку **Принять** в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="bc84c-234">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![Диалоговое окно запроса разрешений](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="bc84c-236">После принятия пользователем запрос разрешений больше не выводится на экран.</span><span class="sxs-lookup"><span data-stu-id="bc84c-236">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

6. <span data-ttu-id="bc84c-237">Надстройка читает данные из OneDrive для бизнеса пользователя, выполнившего вход, и записывает в документ имена из 10 самых популярных файлов и папок.</span><span class="sxs-lookup"><span data-stu-id="bc84c-237">The add-in reads data from the signed-in user's OneDrive for Business and writes the names of the top 10 files and folders to the document.</span></span> <span data-ttu-id="bc84c-238">На следующем рисунке показан пример имен файлов и папок, записанных на лист Excel.</span><span class="sxs-lookup"><span data-stu-id="bc84c-238">The following image shows an example of file and folder names written to an Excel worksheet.</span></span>

    ![Сведения о OneDrive для бизнеса в таблице Excel](../images/sso-onedrive-info-excel.png)

### <a name="outlook"></a><span data-ttu-id="bc84c-240">Outlook</span><span class="sxs-lookup"><span data-stu-id="bc84c-240">Outlook</span></span>

<span data-ttu-id="bc84c-241">Выполните следующие действия, чтобы испытать надстройку Outlook.</span><span class="sxs-lookup"><span data-stu-id="bc84c-241">Complete the following steps to try out an Outlook add-in.</span></span>

1. <span data-ttu-id="bc84c-242">В корневой папке проекта выполните следующую команду, чтобы построить проект и запустить локальный веб-сервер.</span><span class="sxs-lookup"><span data-stu-id="bc84c-242">In the root folder of the project, run the following command to build the project and start the local web server.</span></span>

    > [!NOTE]
    > <span data-ttu-id="bc84c-243">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="bc84c-243">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="bc84c-244">Если вам будет предложено установить сертификат после того, как вы запустите указанную ниже команду, примите предложение установить сертификат, предоставленный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="bc84c-244">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="bc84c-245">Чтобы загрузить неопубликованную надстройку в Outlook, следуйте инструкциями из статьи [Загрузка неопубликованных надстроек Outlook для тестирования](/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span><span class="sxs-lookup"><span data-stu-id="bc84c-245">Follow the instructions in [Sideload Outlook add-ins for testing](/outlook/add-ins/sideload-outlook-add-ins-for-testing) to sideload the add-in in Outlook.</span></span> <span data-ttu-id="bc84c-246">Убедитесь, что вы выполнили вход в Outlook с пользователем, который является участником той же организации Office 365, что и учетная запись администратора Office 365, которую вы использовали для подключения к Azure при [настройке единого входа](sso-quickstart.md#configure-sso) для приложения.</span><span class="sxs-lookup"><span data-stu-id="bc84c-246">Make sure that you're signed in to Outlook with a user that's a member of the same Office 365 organization as the Office 365 administrator account that you used to connect to Azure while [configuring SSO](sso-quickstart.md#configure-sso) for the app.</span></span> <span data-ttu-id="bc84c-247">Благодаря этому будут созданы соответствующие условия для успешного единого входа.</span><span class="sxs-lookup"><span data-stu-id="bc84c-247">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="bc84c-248">В Outlook создайте новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="bc84c-248">In Outlook, compose a new message.</span></span>

4. <span data-ttu-id="bc84c-249">В окне создания сообщения нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="bc84c-249">In the message compose window, choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Кнопка надстройки Outlook](../images/outlook-sso-ribbon-button.png)

5. <span data-ttu-id="bc84c-251">В нижней части области задач нажмите кнопку **прочитать мою службу OneDrive для бизнеса** , чтобы начать процесс единого входа.</span><span class="sxs-lookup"><span data-stu-id="bc84c-251">At the bottom of the task pane, choose the **Read my OneDrive for Business** button to initiate the SSO process.</span></span> 

6. <span data-ttu-id="bc84c-252">Если открывается диалоговое окно, в котором запрашиваются разрешения от имени надстройки, это означает, что единый вход не поддерживается для вашего сценария и надстройка использует альтернативный метод проверки подлинности пользователя.</span><span class="sxs-lookup"><span data-stu-id="bc84c-252">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="bc84c-253">Это может произойти, если администратор клиента не дал согласие на доступ надстройки к Microsoft Graph или если пользователь не вошел в Office с помощью действительной учетной записи Майкрософт или Office 365 (рабочей или учебной учетной записи).</span><span class="sxs-lookup"><span data-stu-id="bc84c-253">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft Account or Office 365 ("Work or School") account.</span></span> <span data-ttu-id="bc84c-254">Чтобы продолжить, нажмите кнопку **Принять** в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="bc84c-254">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![Диалоговое окно запроса разрешений](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="bc84c-256">После принятия пользователем запрос разрешений больше не выводится на экран.</span><span class="sxs-lookup"><span data-stu-id="bc84c-256">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

7. <span data-ttu-id="bc84c-257">Надстройка читает данные из OneDrive для бизнеса пользователя, выполнившего вход, и записывает имена 10 файлов и папок в текст сообщения электронной почты.</span><span class="sxs-lookup"><span data-stu-id="bc84c-257">The add-in reads data from the signed-in user's OneDrive for Business and writes the names of the top 10 files and folders to the body of the email message.</span></span>

    ![Сведения о OneDrive для бизнеса в сообщении Outlook](../images/sso-onedrive-info-outlook.png)

## <a name="next-steps"></a><span data-ttu-id="bc84c-259">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="bc84c-259">Next steps</span></span>

<span data-ttu-id="bc84c-260">Поздравляем, вы успешно настроили функции надстройки с поддержкой единого входа, созданной с помощью генератора Yeoman в [быстром запуске единого входа](sso-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="bc84c-260">Congratulations, you've successfully customized the functionality of the SSO-enabled add-in that you created with the Yeoman generator in the [SSO quick start](sso-quickstart.md).</span></span> <span data-ttu-id="bc84c-261">Дополнительные сведения об этапах настройки единого входа, которые генератор Yeoman выполняет автоматически, и коде, который упрощает процесс единого входа, см. в статье [Создание надстройки Office на платформе Node.js с использованием единого входа](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="bc84c-261">To learn more about SSO configuration steps that the Yeoman generator completed automatically, and the code that facilitates the SSO process, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

## <a name="see-also"></a><span data-ttu-id="bc84c-262">См. также</span><span class="sxs-lookup"><span data-stu-id="bc84c-262">See also</span></span>

- [<span data-ttu-id="bc84c-263">Включение единого входа для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="bc84c-263">Enable single sign-on for Office Add-ins</span></span>](../develop/sso-in-office-add-ins.md)
- [<span data-ttu-id="bc84c-264">Краткое руководство по единому входу (SSO)</span><span class="sxs-lookup"><span data-stu-id="bc84c-264">Single sign-on (SSO) quick start</span></span>](sso-quickstart.md)
- [<span data-ttu-id="bc84c-265">Создание надстройки Office на платформе Node.js с использованием единого входа</span><span class="sxs-lookup"><span data-stu-id="bc84c-265">Create a Node.js Office Add-in that uses single sign-on</span></span>](../develop/create-sso-office-add-ins-nodejs.md)
- [<span data-ttu-id="bc84c-266">Устранение ошибок единого входа</span><span class="sxs-lookup"><span data-stu-id="bc84c-266">Troubleshoot error messages for single sign-on (SSO)</span></span>](../develop/troubleshoot-sso-in-office-add-ins.md)