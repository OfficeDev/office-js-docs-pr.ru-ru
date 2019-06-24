
<span data-ttu-id="66870-101">Выполните следующие действия, чтобы запустить локальный веб-сервер и загрузить неопубликованную надстройку.</span><span class="sxs-lookup"><span data-stu-id="66870-101">Complete the following steps to start the local web server and sideload your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="66870-102">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="66870-102">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="66870-103">Если вам будет предложено установить сертификат после того, как вы запустите одну из указанных ниже команд, примите предложение установить сертификат, предоставленный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="66870-103">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

> [!TIP]
> <span data-ttu-id="66870-104">Если вы тестируете надстройку на компьютере Mac, перед продолжением выполните указанную ниже команду.</span><span class="sxs-lookup"><span data-stu-id="66870-104">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="66870-105">После выполнения этой команды запустится локальный веб-сервер.</span><span class="sxs-lookup"><span data-stu-id="66870-105">When you run this command, the local web server will start.</span></span>
>
> ```command&nbsp;line
> npm run dev-server
> ```

- <span data-ttu-id="66870-106">Чтобы проверить надстройку в Excel, выполните следующую команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="66870-106">To test your add-in in Excel, run the following command in the root directory of your project.</span></span> <span data-ttu-id="66870-107">Когда вы выполните эту команду, запустится локальный веб-сервер (если он еще не запущен) и откроется приложение Excel, в котором будет загружена ваша надстройка.</span><span class="sxs-lookup"><span data-stu-id="66870-107">When you run this command, the local web server will start and Word will open with your add-in loaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

- <span data-ttu-id="66870-108">Чтобы проверить надстройку в Excel в браузере, выполните следующую команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="66870-108">To test your add-in in Excel on a browser, run the following command in the root directory of your project.</span></span> <span data-ttu-id="66870-109">После выполнения этой команды запустится локальный веб-сервер (если он еще не запущен).</span><span class="sxs-lookup"><span data-stu-id="66870-109">When you run this command, the local web server will start.</span></span>

    ```command&nbsp;line
    npm run start:web
    ```

    <span data-ttu-id="66870-110">Чтобы использовать надстройку, откройте новую книгу в Excel в Интернете и затем загрузите неопубликованную надстройку, следуя инструкциям в статье [Загрузка неопубликованных надстроек Office в Office в Интернете](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="66870-110">To use your add-in, open a new document in Word Online and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

