<span data-ttu-id="8eb58-101">Если локальный веб-сервер уже запущен и ваша надстройка уже загружена в Word, перейдите к шагу 2.</span><span class="sxs-lookup"><span data-stu-id="8eb58-101">If the local web server is already running and your add-in is already loaded in Word, proceed to step 2.</span></span> <span data-ttu-id="8eb58-102">В противном случае запустите локальный веб-сервер и Загрузка неопубликованных надстройку:</span><span class="sxs-lookup"><span data-stu-id="8eb58-102">Otherwise, start the local web server and sideload your add-in:</span></span> 

- <span data-ttu-id="8eb58-103">Чтобы протестировать надстройку в Word, выполните следующую команду в корневом каталоге проекта.</span><span class="sxs-lookup"><span data-stu-id="8eb58-103">To test your add-in in Word, run the following command in the root directory of your project.</span></span> <span data-ttu-id="8eb58-104">При этом запустится локальный веб-сервер (если он еще не запущен) и будет открыто приложение Word с загруженной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="8eb58-104">This starts the local web server (if it's not already running) and opens Word with your add-in loaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

- <span data-ttu-id="8eb58-105">Чтобы протестировать надстройку в Word в Интернете, выполните следующую команду в корневом каталоге проекта.</span><span class="sxs-lookup"><span data-stu-id="8eb58-105">To test your add-in in Word on the web, run the following command in the root directory of your project.</span></span> <span data-ttu-id="8eb58-106">При выполнении этой команды запустится локальный веб-сервер (если он еще не запущен).</span><span class="sxs-lookup"><span data-stu-id="8eb58-106">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run start:web
    ```

    <span data-ttu-id="8eb58-107">Чтобы использовать надстройку, откройте новый документ в Word в Интернете и затем Загрузка неопубликованных свою надстройку, следуя инструкциям в статье [Загрузка неопубликованных Office Add-ins in Office in Web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="8eb58-107">To use your add-in, open a new document in Word on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>
