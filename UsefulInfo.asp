<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<br><br>
<h1>Useful Information</h1>
<br>
<h2>MSMQ Queue Information</h2><br>
<table>
<thead>
<tr>
<th nowrap width="250px" align="left">
Queue name
</th>
<th align="left">
Contains information about
</th>
</tr>
</thead>
<tbody>
<tr>
<td nowrap width="300px" align="left">
Enterprise Vault Exchange Mailbox task for <exsrv>1 <number>2 a1 
</td>
<td>
Update Shortcut and Operation Failed 
</td>
</tr>
<tr>
<td>
Enterprise Vault Exchange Mailbox task for <exsrv>1 <number>2 a2
</td>
<td>
Process Item (Explicit Archives). Used for manual archive and when EV cannot communicate with the StorageArchive queue.
</td>
</tr>
<tr>
<td>
Enterprise Vault Exchange Mailbox task for <exsrv>1 <number>2 a3
</td>
<td>
Process Mailbox, Process System (Run Now), Check System, Check Mailbox 
</td>
</tr>
<tr>
<td>
Enterprise Vault Exchange Mailbox task for <exsrv>1 <number>2 a4
</td>
<td>
Only used for retries where Enterprise Vault cannot communicate directly with the StorageArchive queue. 
</td>
</tr>
<tr>
<td>
Enterprise Vault Exchange Mailbox task for <exsrv>1 <number>2 a5
</td>
<td>
Process Mailbox, Process System (Schedule only)
</td>
</tr>
<tr>
<td>
Enterprise Vault Exchange Mailbox task for <exsrv>1 <number>2 a6
</td>
<td>
Synchronization requests. 
</td>
</tr>
<tr>
<td>
Enterprise Vault Exchange Mailbox task for <exsrv>1 <number>2 r1
</td>
<td>
Item Ready, Operation Failed
</td>
</tr>
<tr>
<td>
Enterprise Vault Exchange Mailbox task for <exsrv>1 <number>2 r2
</td>
<td>
Restore Item, Update Basket
</td>
</tr>
<tr>
<td>
Enterprise Vault spool queue 
</td>
<td>
Message Content
</td>
</tr>
<tr>
<td>
Enterprise Vault journal task for <exsrv>1 <number>2 j1
</td>
<td>
Delete Message, Operation Failed
</td>
</tr>
<tr>
<td>
Enterprise Vault journal task for <exsrv>1 <number>2 j2
</td>
<td>
Process Item3
</td>
</tr>
<tr>
<td>
Enterprise Vault journal task for <exsrv>1 <number>2 j3 
</td>
<td>
Instructs the Exchange Journaling task to examine the journal mailbox for new messages.
Up to 250 new messages will be marked as archive pending and a message is placed on queue j2 for each such message. 
</td>
</tr>
<tr>
<td>
Enterprise Vault storage archive for <exsrv>1
</td>
<td>
Store Item
</td>
</tr>
<tr>
<td>
Enterprise Vault storage restore for <exsrv>1
</td>
<td>
Restore an Item
</td>
</tr>
</tbody>
</table>
<br>
<table>
<tr>
<td>
1<exsrv> is the name of the server being processed by the task.<br>  2<number> is a number that uniquely identifies the queue. <br> 3 There is a DWORD registry value, QueueJournalItems, under the Agents key, that controls whether to use J2. The default setting of 1 enables the use of J2; a setting of 0 disables the use of J2.
</td>
</tr>
</table>
<br>
<h2>Further Information</h2>
<br>
<h2>About the Exchange Mailbox task</h2><br>

 

The Exchange Mailbox task uses six queues. Each queue has a different function, which can be monitored to determine the progress of the task.<br><br>

 

<h3>Update Shortcut</h3><br>

Informs the Exchange Mailbox task to turn an archive pending item into a shortcut. It occurs after a message has been stored by the Storage service, and backed up.<br><br>

 

<h3>Operation Failed</h3><br>

Informs the Exchange Mailbox task that an error has occurred and it should change the message from archive pending back into a message. The message will be reprocessed later. This message is sent if an error occurs during archiving and storage.<br><br>

 

<h3>Process Item</h3><br>

Asks the Exchange Mailbox task to archive a specific message from the exchange server to the Storage service. The item in exchange will be turned into a shortcut when the storage returns an Update Shortcut message. Process Item messages are produced by a user explicitly archiving a message (placed on A2) or during a Process Mailbox (in which case the message is placed on queue A4 if from a scheduled process mailbox or on queue A2 for an immediate Run Now archive).<br><br>

 

<h3>Process Mailbox</h3><br>

Asks the Exchange Mailbox task to examine a mailbox, finding any messages, which match the archiving criteria. These messages are then turned into archive pending, and a message is placed on the process item queue for each message to be archived.  The process item queue will be the StorageArchive queue of the Storage service if the process mailbox is generated from a scheduled archive, or is as a result of an administrator Run Now.<br><br>

All Explicit Archives, that is, individual users manually archiving items, result in a Process Item message being placed on the A2 queue.
<br><br>
 

<h3>Process System</h3><br>

Asks the Exchange Mailbox task to determine which mailboxes on the exchange server are eligible for archiving. The Exchange Mailbox task reads the list of all enabled mailboxes and sends a Process Mailbox message (on the same queue) for each mailbox eligible.<br><br>

Process System message is placed immediately on queue A3 if the administrator selects Run Now from the task properties, or it will be placed on queue A5 at the start of a scheduled archive period (provided that there are no other process system messages already waiting to be done).<br><br>

 

<h3>Check System</h3><br>

Asks the Exchange Mailbox task to determine which mailboxes on the exchange server are eligible for archiving. The Exchange Mailbox task reads the list of all mailboxes, which have been used with Enterprise Vault, and places a Check Mailbox message on the queue for each mailbox eligible.<br><br>

This message will only be placed on queue A3 at the start time specified in the site properties.<br><br>

 

<h3>Check Mailbox</h3><br>

Asks the Exchange Mailbox task to examine a mailbox, finding any shortcuts, which match the expiry or shortcut deletion criteria. These shortcuts are then deleted from the mailbox.<br><br>

Check Mailbox messages are only placed on queue A3.<br><br>

 

 

<h3>Synchronize System</h3><br>

Synchronization can take a large amount of time so synchronization is multi-threaded, using the agents queues.<br><br>

Synchronization requests for an Exchange Mailbox task are placed on the A6 queue. When synchronization is run, a Synchronize System request is placed on this queue and this generates a Synchronize Mailbox request for each mailbox that needs to be synchronized. Having multiple Synchronize Mailbox requests means that the requests can be serviced by multiple threads.<br><br>

The A6 queue is processed at all times but is always the lowest priority task. This means that scheduled background archives always take precedence over a synchronize.<br><br>

 

<h3>Notes</h3><br>

Each queue has a suffix of A<priority number>, where A1 is the highest priority and A5 is the lowest priority. The message queues are treated as FIFO (First In, First Out), and new messages are always added to the end of the queue.<br><br>

 

The Exchange Mailbox task processes the queues in order of priority. The task scans through each queue, starting with the highest priority. If it finds a message on a queue, it processes the message, then starts the scan again from the highest priority queue. Therefore queues A2 through A5 will not be processed until queue A1 is empty.<br><br>

However queue A5 is a special queue that is used only by the archiving schedule. The Exchange Mailbox task processes messages on the A5 queue only during a scheduled archive period. Outside the scheduled periods messages on these queues are ignored.<br><br>

 

Using performance monitor, you can monitor the changes in the queues to assess the progress of the task.<br><br>

 

For example, at the start of a scheduled period, the number of messages on queue A5 rises (to the number of enabled mailboxes on the exchange server). These are Process Mailbox messages. The Exchange Mailbox task will take the first message off queue A5 and find all the eligible messages in the mailbox and change them to archive pending. A Process Item message is then placed on the StorageArchive of the Storage service for each message to be archived.<br><br>

After the vault store has been backed up, Update Shortcut messages will be placed on queue A1 — which will be processed immediately because the queue has a higher priority.<br><br>

 

Queue A3 performs the same function as queue A5, but for an immediate process system. This queue also performs shortcut expiry and deletion. Explicit user archives from the Outlook client extension are placed on queue A2.<br><br>

 

Queue A5 will only be processed during a scheduled period, but queues A1-A3 will always be processed. If a queue is not being processed (the number of messages is not changing) for more than 10 minutes, and there are no messages in a higher priority queue, then there may be a problem with the task. Check the Enterprise Vault Event Log on the Exchange Mailbox task computer for any additional information.<br><br>

 

Monitoring queue A1 will indicate that a backup has correctly updated shortcuts, but if A1 is being used during normal use (before a backup), then it may indicate a problem with tasks. Check the Enterprise Vault Event Log for errors.<br><br>

 

<h2>About the Storage service</h2><br>

The Storage service uses two queues. Each queue has a different function, which can be monitored to determine the progress of the service.<br><br>

 

<h3>Store Item</h3><br>

The Exchange Mailbox task will place a compressed email message on the Storage Archive queue, to be stored in an archive. If the compressed message is larger than 4MB then it will be divided into 4MB chunks (each labeled with a part number, for example Part 1 of 5). The message is reconstructed by the Storage service before storing.<br><br>

The Exchange Mailbox task will place all e-mails to be stored onto the appropriate Storage service archive queue (multiple Storage services may be configured, therefore the Exchange Mailbox task must select the correct Storage service for the vault store in which the archive resides).<br><br>

 

<h3>Restore an Item</h3><br>

The Exchange Mailbox task will place a message onto the storage restore queue, requesting an item to be restored from an archive. When the Storage service has located the item, it places it onto the Storage Spool queue, and notifies the Exchange Mailbox task on queue R2.<br><br>

 

<h3>Notes</h3><br>

Monitoring the storage archive queue will indicate the Storage service is processing items. If the number of items in this queue does not change for at least 30 minutes, then there is likely to be a problem. Check the Enterprise Vault Event Log on the storage computer for any errors, and look at the task list process storagearchive.exe. If the process is at 0% CPU then it has stopped doing any work. To correct the problem, restart the service.<br><br>

Monitoring the restore queue will indicate the number of restores required by the users. Again if the number of items on the queue does not change then there is likely to be a problem.<br><br>

 

<h3>About retrieval</h3><br>

Retrieval is carried out by the Exchange Mailbox task, using three queues. Each queue has a different function, which can be monitored to determine the progress.<br><br>

 

<h3>Restore Item</h3><br>

This message is a request to restore an item from the Storage service back into Exchange Server. The Exchange Mailbox task prompts for the message from the Storage service and places it into the mailbox. Messages are placed on this queue from both the user extension and the Web page restore functions.<br><br>

 

<h3>Operation Failed</h3><br>

This message informs the Exchange Mailbox task there was a problem restoring the message. If the retrieval was started from the Web application, then the Exchange Mailbox task will update the basket to indicate the item was not restored.<br><br>

 

<h3>Update Basket</h3><br>

This message informs the Exchange Mailbox task to update a Web basket with the successful restoration of an item.<br><br>

 

<h3>Item Ready</h3><br>

This message informs the Exchange Mailbox task that a previously requested message is now available on the storage spool queue. The Exchange Mailbox task will collect the message from the storage spool queue and place it into the mailbox. These messages are generated by the Storage service as required.<br><br>

 

<h3>Storage Spool</h3><br>

The messages on this queue are items restored from the Storage service. The Exchange Mailbox task will read the messages as it processes the queue R2.<br><br>

 

<h3>Notes</h3><br>

Each queue has a suffix of R<priority number>, where R1 is the highest priority and R2 is the least important. The message queues are treated as first in, first out (FIFO), and new messages are always added to the end of the queue.<br><br>

 

The Exchange Mailbox task will process the queues in order of priority. The task scans through each queue, starting with the highest priority. If it finds a message on a queue, it will process the message, then start the scan again from the highest priority queue. Therefore if there are messages on queue R1, then queue R2 will not be processed until queue R1 is empty.<br><br>

 

If a queue is not being processed (the number of messages is not changing) for more than 10 minutes, and there are no messages in a higher priority queue, then there may be a problem with the task. Check the Enterprise Vault Event Log on the Exchange Mailbox task computer for any additional information.<br><br>

 

<h2>About the Exchange Journaling task</h2><br>

An Exchange Journaling task uses four queues. Each queue has a different function, which can be monitored to determine the progress of the task.<br><br>

 

<h3>Delete Message</h3><br>

Informs the Exchange Journaling task to delete an archive pending item from the journal mailbox. It occurs after a message has been stored by the Storage service, and backed-up.<br><br>

 

<h3>Operation Failed</h3><br>

Informs the Exchange Journaling task that an error has occurred and it should change the message from archive pending back into a message. The message will be reprocessed later. This message is sent if an error occurs during archiving and storage.<br><br>

 

<h3>Process Mailbox</h3><br>

Asks the Exchange Journaling task to examine the journal mailbox, finding any messages that have arrived. Up to 250 new messages will be turned into archive pending, and a message is placed on the StorageArchive queue of the Storage service for each message to be archived.<br><br>

The process mailbox message is issued every minute onto queue J3.<br><br>

 

<h3>Synchronize System</h3><br>

Synchronization is multi-threaded, using the agents queues.<br><br>

Synchronization requests for an Exchange Mailbox task are placed on a new J4 queue. When synchronization is run, a Synchronize System request is placed on this queue and this generates a Synchronize Mailbox request for each mailbox that needs to be synchronized. Having multiple Synchronize Mailbox requests means that the requests can be serviced by multiple threads.<br><br>

The J4 queue is processed at all times but is always the lowest priority task. This means that scheduled background archives always take precedence over a synchronize.<br><br>

 

<h3>Notes</h3><br>

Each queue has a suffix of J<priority number>, where J1 is the highest priority and J3 is the least. The message queues are treated as FIFO, and new messages are always added to the end of the queue.<br><br>

 

A Exchange Journaling task processes the queues in order of priority. The task scans through each queue, starting with the highest priority. If it finds a message on a queue, it will process it, then start the scan again from the highest priority queue. Therefore, if there are messages on queue J1, queue J2 and J3 will not be processed until queue J1 is empty.<br><br>

 

Monitoring queue J1 will indicate that a vault store backup is correctly deleting the messages, but if J1 is being used during normal use (before a backup), then it may indicate a problem with tasks. Check the Enterprise Vault Event Log for errors.<br><br>

 

Monitoring queue J3 will indicate that at least every minute a process mailbox message is on the queue (a new message will only be added if the queue is empty). There should never be more than 1 message on this queue. The message should appear on the queue, then disappear as soon as queue J1 is clear — any new messages in the journal mailbox will be processed.<br><br>

 

If a queue is not being processed (the number of messages is not changing) for more than 10 minutes, and there are no messages in a higher priority queue, then there may be a problem with the tasks. Check the Enterprise Vault Event Log on the Exchange Journaling task computer for any additional information. <br><br>
<!--#include file="footer.asp"-->