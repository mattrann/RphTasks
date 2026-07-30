@@
-    ; JS payload: find any element containing the blueprint text and click it\n-    JSPayload := "(function(){var txt='RX Not Dispensable';var els=document.querySelectorAll('button,div,span,a');for[...]
-
-    ; Try to click blueprint and verify expected task text appears\n-    ClickBlueprintAndVerify(JSPayload, "RX Not Dispensable", false)
+    ; JS payload: click the RX Not Dispensable blueprint and open the expected task UI
+    JSPayload := "(function(){var txt='RX Not Dispensable';var els=document.querySelectorAll('button,div[role=button],a');for(var i=0;i<els.length;i++){try{if(els[i].innerText&&els[i].innerText.trim().indexOf(txt)!=-1){els[i].click();return 'clicked';}}catch(e){} } return 'notfound';})()"
+
+    ; Click blueprint and verify resulting page shows the task title
+    ClickBlueprintAndVerify(JSPayload, "RX Not Dispensable", false)
@@
-    ; Pre-fill the Instructions box with the blueprint template\n-    template := "Instructions to CCC Agent:`n`nCustomer name: {{ticket.requester.name}}`nDate of Birth: {{ticket.requester.custom_f[...]
-    FindAndPaste("Instructions to CCC Agent:", template)
+    ; Pre-fill the Instructions box using the same template as the blueprint (matches NDC OOS style)
+    template := "Instructions to CCC Agent:`n`nCustomer name: {{ticket.requester.name}}`nDate of Birth: {{ticket.requester.custom_fields.birthdate}}`nFacility: {{ticket.requester.custom_fields.facility_name}}`n`nShipment in question (please provide a shipment link):`n`nNot Dispensable Prescription #:`n`nMedication Name(s):"
+    FindAndPaste("Instructions to CCC Agent:", template)
*** End Patch