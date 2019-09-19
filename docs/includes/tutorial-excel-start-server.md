<span data-ttu-id="6fcc6-101">Если локальный веб-сервер уже запущен и ваша надстройка уже загружена в Excel, перейдите к шагу 2.</span><span class="sxs-lookup"><span data-stu-id="6fcc6-101">If the local web server is already running and your add-in is already loaded in Excel, proceed to step 2.</span></span> <span data-ttu-id="6fcc6-102">В противном случае запустите локальный веб-сервер и Загрузка неопубликованных надстройку:</span><span class="sxs-lookup"><span data-stu-id="6fcc6-102">Otherwise, start the local web server and sideload your add-in:</span></span> 

- <span data-ttu-id="6fcc6-103">Чтобы протестировать надстройку в Excel, выполните следующую команду в корневом каталоге проекта.</span><span class="sxs-lookup"><span data-stu-id="6fcc6-103">To test your add-in in Excel, run the following command in the root directory of your project.</span></span> <span data-ttu-id="6fcc6-104">При этом запустится локальный веб-сервер (если он еще не запущен) и откроется Excel с загруженной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="6fcc6-104">This starts the local web server (if it's not already running) and opens Excel with your add-in loaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

- <span data-ttu-id="6fcc6-105">Чтобы протестировать надстройку в Excel в Интернете, выполните следующую команду в корневом каталоге проекта.</span><span class="sxs-lookup"><span data-stu-id="6fcc6-105">To test your add-in in Excel on the web, run the following command in the root directory of your project.</span></span> <span data-ttu-id="6fcc6-106">При выполнении этой команды запустится локальный веб-сервер (если он еще не запущен).</span><span class="sxs-lookup"><span data-stu-id="6fcc6-106">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run start:web
    ```

    <span data-ttu-id="6fcc6-107">Чтобы использовать надстройку, откройте новый документ в Excel в Интернете и затем Загрузка неопубликованных свою надстройку, следуя инструкциям в статье [Загрузка неопубликованных Office Add-ins in Office in Web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="6fcc6-107">To use your add-in, open a new document in Excel on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>
