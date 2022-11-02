---
title: Размещение надстройки Office в Microsoft Azure | Документация Майкрософт
description: Сведения о развертывании веб-приложения надстройки в Azure и загрузке неопубликованной надстройки для тестирования в клиентском приложении Office.
ms.date: 07/07/2020
ms.localizationpriority: medium
ms.openlocfilehash: d80dafeab272f1649d9487f284e44e41cf34c1ef
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810031"
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a>Размещение надстройки Office в Microsoft Azure

Простейшая надстройка Office состоит из XML-файла манифеста и HTML-страницы. XML-файл манифеста описывает характеристики надстройки, такие как ее имя, то, в каких классических клиентах Office она может запускаться, а также URL-адрес HTML-страницы надстройки. HTML-страница содержится в веб-приложении, с которым пользователь взаимодействует, когда устанавливает и запускает надстройку в клиентском приложении Office. Вы можете разместить веб-приложение надстройки Office на любой платформе веб-хостинга, включая Azure.

В этой статье рассказывается, как развернуть веб-приложение надстройки в Azure и [загрузить неопубликованную надстройку](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) для тестирования в клиентском приложении Office.

## <a name="prerequisites"></a>Предварительные требования 

1. Установите [Visual Studio 2019](https://www.visualstudio.com/downloads) и не забудьте включить рабочую нагрузку **Разработка для Azure**.

    > [!NOTE]
    > Если Visual Studio 2019 уже установлен, убедитесь, что рабочая нагрузка **Разработка для Azure** установлена, [используя установщик Visual Studio](/visualstudio/install/modify-visual-studio). 

2. Установите Office.

    > [!NOTE]
    > Если у вас еще нет Office, можете [оформить бесплатную пробную подписку на 1 месяц](https://products.office.com/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).

3. Подпишитесь на Azure.

    > [!NOTE]
    > Если у вас еще нет подписки на Azure, вы можете [получить ее в рамках своей подписки на Visual Studio](https://azure.microsoft.com/pricing/member-offers/visual-studio-subscriptions/) или [зарегистрировать бесплатную учетную запись](https://azure.microsoft.com/pricing/free-trial). 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a>Шаг 1. Создание общей папки для размещения XML-файла манифеста надстройки

1. Откройте проводник на своем компьютере разработчика.

2. Щелкните диск C: правой кнопкой мыши и выберите пункты **Создать** > **Папку**.

3. Назовите новую папку AddinManifests.

4. Щелкните папку AddinManifests правой кнопкой мыши и выберите пункты **Общий доступ** > **Конкретные пользователи...**.

5. В окне **Общий доступ к файлам** щелкните стрелку раскрывающегося списка и выберите **Все** > **Добавить** > **Общий доступ**.

> [!NOTE]
> In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file. In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](/office/dev/store/submit-to-appsource-via-partner-center).

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a>Шаг 2. Добавление общей папки в доверенный каталог надстроек

1. Запустите Word и создайте документ.

    > [!NOTE]
    > В этом примере используется Word, но вы можете использовать любое приложение Office, поддерживающее надстройки Office, например Excel, Outlook, PowerPoint или Project.

2. Щелкните **Файл** > **Параметры**.

3. В диалоговом окне **Параметры Word** щелкните **Центр управления безопасностью**, а затем — **Параметры центра управления безопасностью**.

4. In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**. Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**. 

5. Установите флажок **Показывать в меню**.

    > [!NOTE]
    > Когда XML-файл манифеста надстройки хранится в доверенном каталоге веб-надстроек, надстройка отображается в разделе **Общая папка** в диалоговом окне **Надстройки Office** (**Вставка** > **Мои надстройки**).

6. Закройте Word.

## <a name="step-3-create-a-web-app-in-azure-using-the-azure-portal"></a>Шаг 3. Создание веб-приложения в Azure с помощью портала Azure

Чтобы создать веб-приложение с помощью портала Azure, выполните указанные ниже действия.

1. Войдите в систему на [портале Azure](https://portal.azure.com/), используя свои учетные данные Azure.

2. В разделе **Службы Azure** выберите **Веб-приложения**.

3. На странице **Служба приложений** выберите **Добавить**. Чтобы добавить эти сведения, выполните указанные ниже действия.

      - Выберите **подписку**, которую необходимо использовать для создания сайта.

      - Выберите **группу ресурсов** для своего сайта. Если вы создадите группу, вам потребуется присвоить ей имя.

      - Введите уникальное **имя приложения** для своего сайта. Azure проверит уникальность имени сайта в домене azureweb apps.net.

      - Укажите, следует ли выполнить публикацию с помощью кода или контейнера Docker.

      - Укажите **Стек среды выполнения**.

      - Выберите **операционную систему** для своего сайта.

      - Выберите **Регион**.

      - Выберите **план службы приложений**, который необходимо использовать для создания этого сайта.

      - Нажмите кнопку **Создать**.

4. На следующей странице вы узнаете о том, как выполняется развертывание и когда оно завершится. После завершения развертывания выберите пункт **Перейти к ресурсу**.  

5. В разделе **Обзор** выберите URL-адрес, который отображается в пункте **URL**. Откроется браузер, и в нем отобразится веб-страница с сообщением "Ваша служба приложений готова к работе".

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Веб-сайты Azure автоматически предоставляют конечную точку HTTPS.

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a>Шаг 4. Создание надстройки Office в Visual Studio

1. Запустите Visual Studio от имени администратора.

2. Выберите **Создание нового проекта**.

3. Используя поле поиска, введите **надстройка**.

4. Выберите пункт **Веб-надстройка Word** в качестве типа проекта, а затем нажмите кнопку **Далее**, чтобы принять параметры, используемые по умолчанию.

Visual Studio создаст базовую надстройку Word, которую вы можете опубликовать в том виде, в котором она есть, не внося изменений в ее веб-проект. Чтобы создать надстройку для другого приложения Office, например Excel, повторите шаги и выберите тип проекта с нужным приложением Office.

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a>Действие 5. Публикация веб-приложения надстройки Office в Azure

1. Не закрывая проект вашей надстройки в Visual Studio, разверните узел решения в **Обозревателе решений**, затем выберите **Служба приложений**.

2. Right-click the web project and then choose **Publish**. The web project contains Office Add-in web app files so this is the project that you publish to Azure.

3. На вкладке **Публикация** выполните указанные ниже действия.

      - Выберите пункт **Служба приложений Microsoft Azure**.

      - Щелкните **Выбрать существующую**.

      - Щелкните **Опубликовать**.

4. Visual Studio publishes the web project for your Office Add-in to your Azure web app. When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created." This is the current default page for the web app.

5. Скопируйте корневой URL-адрес (например, `https://YourDomain.azurewebsites.net`); он понадобится при изменении файла манифеста надстройки далее в этой статье.

## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a>Шаг 6. Редактирование и развертывание XML-файла манифеста надстройки

1. В Visual Studio (с примером надстройки Office, открытом в **обозревателе решений**) разверните решение так, чтобы отображались оба проекта.

2. Разверните проект надстройки Office (например, WordWebAddIn), щелкните правой кнопкой мыши папку манифеста, а затем нажмите кнопку **Открыть**. Откроется XML-файл манифеста надстройки.

3. В XML-файле манифеста найдите и замените все фрагменты ~remoteAppUrl URL-адресом корня веб-приложения надстройки в Azure. Это URL-адрес, скопированный ранее после публикации веб-приложения надстройки в Azure (например, `https://YourDomain.azurewebsites.net`).

4. Щелкните **Файл** и выберите пункт **Сохранить все**. Затем скопируйте XML-файл манифеста надстройки (например, WordWebAddIn.xml).

5. С помощью программы **Проводник** откройте сетевой файловый ресурс, который вы создали в [действии 1 "Создание общей папки"](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) и вставьте файл манифеста в папку.

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a>Шаг 7. Вставка и запуск надстройки в клиентском приложении Office

1. Запустите Word и создайте документ.

2. На ленте щелкните **Вставка** > **Мои надстройки**.

3. In the **Office Add-ins** dialog box, choose **SHARED FOLDER**. Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box. You should see an icon for your sample add-in.

4. Choose the icon for your add-in and then choose **Add**. A **Show Taskpane** button for your add-in is added to the ribbon.

5. On the ribbon of the **Home** tab, choose the **Show Taskpane** button. The add-in opens in a task pane to the right of the current document.

6. Verify that the add-in works by selecting some text in the document and choosing the **Highlight!** button in the task pane.

## <a name="see-also"></a>См. также

- [Публикация надстройки Office](../publish/publish.md)
- [Публикация надстройки с помощью Visual Studio](../publish/package-your-add-in-using-visual-studio.md)
