/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */


  Office.onReady(info => {
    if (info.host === Office.HostType.Excel) {
      document.getElementById("app-body").style.display = "block";
      document.getElementById("authorize").onclick = handleAuthorizeButtonClick;
    }
  });
  function handleAuthorizeButtonClick() {
    const authUrlSelect = document.getElementById("authUrl");
    const selectedValue = authUrlSelect.value;
    console.log(selectedValue);
  }
  