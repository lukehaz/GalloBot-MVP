const notificationTemplate = require("./adaptiveCards/notification-default.json"); // Pass in the adaptive card template
const { notificationApp } = require("./internal/initialize");
const { AdaptiveCards } = require("@microsoft/adaptivecards-tools");
const { TeamsBot } = require("./teamsBot");
const restify = require("restify");



// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());
server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nApp Started, ${server.name} listening to ${server.url}`);
});

// HTTP trigger to send notification. You need to add authentication / authorization for this API. Refer https://aka.ms/teamsfx-notification for more details.
server.post(
  "/api/notification",
  restify.plugins.queryParser(),
  restify.plugins.bodyParser(), // Add more parsers if needed
  async (req, res) => {
    const pageSize = 100;
    let continuationToken = undefined;

    let payloadData = req.body; // Use the JSON from the request body

    // Fallback to default payload if none provided
    if (!payloadData || typeof payloadData !== 'object') {
      payloadData = {
        gpm: "-896 gal",
        tank: "2001",
        workOrder: "24-120821",
        tankNetwork: "default test",
        duration: 3600,
        cesUrl: "https://www.google.com/"
      };
    }



    do {
      const pagedData = await notificationApp.notification.getPagedInstallations(
        pageSize,
        continuationToken
      );
      const installations = pagedData.data;
      continuationToken = pagedData.continuationToken;

      for (const target of installations) {
        
        if (target.type === "Channel") {
          const members = await target.members();

          // Below is to send to channel
          // Await target.sendMessage("This is a message to channel: " + members.length + " members.");
         
         
          // Below is how to send a message to a specific memeber using email
          /*
          for (const member of members) {
            if (member.account.email === "IsaiahL@5ml88b.onmicrosoft.com") {
            await member.sendMessage("This is a message to Member: " + member.account.email);
            }
          } */
        }
        

        await target.sendAdaptiveCard( //target is channel, sends to channel
          AdaptiveCards.declare(notificationTemplate).render(payloadData));


        
        /****** To distinguish different target types ******/
        /** "Channel" means this bot is installed to a Team (default to notify General channel)
        if (target.type === NotificationTargetType.Channel) {
          // Directly notify the Team (to the default General channel)
          await target.sendAdaptiveCard(...);

          // List all channels in the Team then notify each channel
          const channels = await target.channels();
          for (const channel of channels) {
            await channel.sendAdaptiveCard(...);
          }

          // List all members in the Team then notify each member
          const pageSize = 100;
          let continuationToken = undefined;
          do {
            const pagedData = await target.getPagedMembers(pageSize, continuationToken);
            const members = pagedData.data;
            continuationToken = pagedData.continuationToken;

            for (const member of members) {
              await member.sendAdaptiveCard(...);
            }
          } while (continuationToken);
        }
        **/

        /** "Group" means this bot is installed to a Group Chat
        if (target.type === NotificationTargetType.Group) {
          // Directly notify the Group Chat
          await target.sendAdaptiveCard(...);

          // List all members in the Group Chat then notify each member
          const pageSize = 100;
          let continuationToken = undefined;
          do {
            const pagedData = await target.getPagedMembers(pageSize, continuationToken);
            const members = pagedData.data;
            continuationToken = pagedData.continuationToken;

            for (const member of members) {
              await member.sendAdaptiveCard(...);
            }
          } while (continuationToken);
        }
        **/

        /** "Person" means this bot is installed as a Personal app
        if (target.type === NotificationTargetType.Person) {
          // Directly notify the individual person
          await target.sendAdaptiveCard(...);
        }
        **/
      }
    } while (continuationToken);

    /** You can also find someone and notify the individual person
    const member = await notificationApp.notification.findMember(
      async (m) => m.account.email === "someone@contoso.com"
    );
    await member?.sendAdaptiveCard(...);
    **/

    /** Or find multiple people and notify them
    const members = await notificationApp.notification.findAllMembers(
      async (m) => m.account.email?.startsWith("test")
    );
    for (const member of members) {
      await member.sendAdaptiveCard(...);
    }
    **/

    res.json({status: "Notifaction Sent"});
  }
);

// Bot Framework message handler.
const teamsBot = new TeamsBot();
server.post("/api/messages", async (req, res) => {
  await notificationApp.requestHandler(req, res, async (context) => {
    await teamsBot.run(context);
  });
});
