const _ = LodashGS.load();

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Slack")
    .addItem("Authorize", "_onAuthorize")
    .addItem("Reset Authorization", "_onReset")
    .addItem("Check Authorization", "_onCheck")
    .addToUi();
}

function _onAuthorize() {
  const authorizationUrl = Slack.getOrCreate("ipdx").getAuthorizationUrl();
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(
      `<script>function authorize() { window.open("${authorizationUrl}", '_blank').focus(); google.script.host.close(); }</script><button onclick="authorize()">Authorize</button>`
    ),
    "Authorize"
  );
}

function _onReset() {
  Slack.getOrCreate("ipdx").reset();
}

function _onCheck() {
  const hasAccess = Slack.getOrCreate("ipdx").hasAccess();
  SpreadsheetApp.getUi().alert(hasAccess ? "Authorized" : "Not Authorized");
}

function getChannels() {
  const IPDX = Slack.getOrCreate("ipdx");
  const publicChannels = IPDX.getPaginated(
    "conversations.list",
    "types",
    "public_channel",
    "exclude_archived",
    "true"
  );
  const privateChannels = IPDX.getPaginated(
    "conversations.list",
    "types",
    "public_channel",
    "exclude_archived",
    "true"
  );
  let channels = [...publicChannels, ...privateChannels];
  channels = flatMap(channels, "channels");
  channels = mapReduce(channels, "id", "name", "is_private", "num_members");
  return toArray(channels);
}

function getLastMessageTimestamp(channelId: string) {
  const IPDX = Slack.getOrCreate("ipdx");
  let messages = IPDX.get("conversations.history", "channel", channelId);
  messages = flatMap(messages, "messages");
  messages = mapReduce(messages, "ts", "type");
  return messages.find((message: any) => message.type === "message")?.ts;
}

function getChannelsWithLastMessageTimestamp() {
  const IPDX = Slack.getOrCreate("ipdx");
  const publicChannels = IPDX.getPaginated(
    "conversations.list",
    "types",
    "public_channel",
    "exclude_archived",
    "true"
  );
  const privateChannels = IPDX.getPaginated(
    "conversations.list",
    "types",
    "public_channel",
    "exclude_archived",
    "true"
  );
  let channels = [...publicChannels, ...privateChannels];
  channels = flatMap(channels, "channels");
  channels = mapReduce(
    channels,
    "id",
    "name",
    "is_private",
    "num_members",
    "is_member"
  );

  channels = channels.map((channel) => {
    if (channel.is_member) {
      channel.last_message_ts = getLastMessageTimestamp(channel.id);
      return channel;
    } else {
      return channel;
    }
  });

  return toArray(channels);
}

function flatMap(data: any[], ...args: any[]): any[] {
  if (!_.isArray(data)) {
    data = [data];
  }
  return _.flatMap(data, (row) => {
    return args.flatMap((arg) => row[arg]);
  });
}

function mapReduce(data: any[], ...args: any[]): any[] {
  return _.map(data, (row) => {
    return _.reduce(
      args,
      (acc, arg) => {
        acc[arg] = row[arg];
        return acc;
      },
      {}
    );
  });
}

function toArray(data: any) {
  const keys = _.uniq(_.flatMap(data, (row) => Object.keys(row)));
  const rows = [keys];
  for (const row of data) {
    const values = [];
    for (const key of keys) {
      values.push(row[key]);
    }
    rows.push(values);
  }
  return rows;
}
