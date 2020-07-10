---
title: Отладка надстроек в Office в Интернете
description: Сведения о том, как тестировать и отлаживать надстройки в Office в Интернете.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: f7ef3fa3d6389629e28b428b9bdbe3b128896b1f
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094493"
---
# <a name="debug-add-ins-in-office-on-the-web"></a>Отладка надстроек в Office в Интернете

Вы можете создавать надстройки и выполнять их отладку на компьютере, на котором нет Windows или классического клиента Office (например, если вы создаете надстройку на компьютере Mac). В этой статье описано, как использовать Office а Интернете для тестирования и отладки надстроек. 

## <a name="prerequisites"></a>Необходимые компоненты

Чтобы приступить к работе, выполните указанные ниже действия.

- Получите учетную запись разработчика Microsoft 365, если у вас еще нет доступа к сайту SharePoint.

  > [!NOTE]
  > To get a free, 90-day renewable Microsoft 365 developer subscription, join our [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program). See the [Microsoft 365 developer program documentation](/office/developer-program/office-365-developer-program) for step-by-step instructions about how to join the Microsoft 365 developer program and configure your subscription.

- Set up an app catalog on SharePoint Online. An app catalog is a dedicated site collection in SharePoint Online that hosts document libraries for Office Add-ins. If you have your own SharePoint site, you can set up an app catalog document library. For more information, see [Publish task pane and content add-ins to an app catalog on SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).


## <a name="debug-your-add-in-from-excel-or-word-on-the-web"></a>Отладка надстройки в Excel и Word в Интернете

Для отладки надстройки с помощью Office в Интернете выполните указанные ниже действия.

1. Разверните надстройку на сервере, поддерживающем SSL.

    > [!NOTE]
    > Рекомендуем использовать [генератор Yeoman](https://github.com/OfficeDev/generator-office) для создания и размещения надстройки.

2. In your [add-in manifest file](../develop/add-in-manifests.md), update the **SourceLocation** element value to include an absolute, rather than a relative, URI. For example:

    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```

3. Отправьте манифест в библиотеку надстроек Office в каталоге приложений SharePoint.

4. Запустите Excel или Word в Интернете из средства запуска приложений в Microsoft 365 и откройте новый документ.

5. Чтобы добавить вашу надстройку и протестировать ее в приложении, на вкладке "Вставка" выберите **Мои надстройки** или **Надстройки Office**.

6. Выполните отладку надстройки в удобном для вас браузерном отладчике.

## <a name="potential-issues"></a>Возможные проблемы

Ниже указаны некоторые проблемы, которые могут возникнуть при отладке.

- Причиной некоторых отображаемых ошибок JavaScript может быть Office в Интернете.

- В браузере может появиться сообщение об ошибке, связанной с недопустимым сертификатом, которое необходимо обойти. Этот процесс зависит от браузера, и пользовательские интерфейсы различных браузеров, предназначенные для его выполнения, периодически изменяются. Инструкции можно найти в справке браузера или выполнить поиск в Интернете. (Например, выполните поиск по запросу "Предупреждение Microsoft Edge о недействительном сертификате".) В большинстве браузеров на странице предупреждения содержится ссылка, позволяющая перейти к странице надстройки. Например, в Microsoft Edge есть ссылка "Перейти на веб-страницу (не рекомендуется)". При этом, как правило, вам потребуется переходить по этой ссылке при каждой перезагрузке надстройки. Сведения о более долговечных способах обхода см. в предложенных разделах справки.

- Если вы задаете точки останова в коде, Office в Интернете может вывести сообщение об ошибке, свидетельствующее о том, что не удается выполнить сохранение.

## <a name="see-also"></a>См. также

- [Рекомендации по разработке надстроек Office](../concepts/add-in-development-best-practices.md)
- [Политики проверки AppSource](/legal/marketplace/certification-policies)  
- [Создание эффективных приложений и надстроек AppSource](/office/dev/store/create-effective-office-store-listings)  
- [Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office](testing-and-troubleshooting.md)
