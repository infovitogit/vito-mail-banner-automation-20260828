"use strict";

var assert = require("assert");
var banner = require("../banner-runtime.js");

function withActiveCampaign(testFn) {
  var originalNow = Date.now;
  Date.now = function () {
    return Date.parse("2026-08-28T12:00:00+03:00");
  };
  try {
    testFn();
  } finally {
    Date.now = originalNow;
  }
}

withActiveCampaign(function () {
  var newMail = "<html><body><div id=\"signature\">Personal signature</div></body></html>";
  var result = banner.insertBannerIntoFullBody(newMail);
  assert.strictEqual(result.changed, true);
  assert.ok(result.html.indexOf("Personal signature") < result.html.indexOf("data-vito-mail-banner"));
  assert.ok(result.html.indexOf("data-vito-mail-banner") < result.html.indexOf("</body>"));
});

withActiveCampaign(function () {
  var reply = "<div>Personal signature</div><a name=\"_MailOriginal\"></a><div>Old message</div>";
  var result = banner.insertBannerIntoFullBody(reply);
  assert.strictEqual(result.changed, true);
  assert.ok(result.html.indexOf("Personal signature") < result.html.indexOf("data-vito-mail-banner"));
  assert.ok(result.html.indexOf("data-vito-mail-banner") < result.html.indexOf("_MailOriginal"));
});

withActiveCampaign(function () {
  var previousBanner = [
    "<div>Current personal signature</div>",
    "<div id=\"divRplyFwdMsg\">",
    "<div data-vito-mail-banner=\"vito-gastech-bangkok-2026\">Previous banner</div>",
    "</div>"
  ].join("");
  var result = banner.insertBannerIntoFullBody(previousBanner);
  assert.strictEqual(result.changed, true, "A previous quoted banner must not suppress the current banner.");
});

withActiveCampaign(function () {
  var alreadyPresent = "<div>Signature</div>" + banner.buildBannerHtml();
  var result = banner.insertBannerIntoFullBody(alreadyPresent);
  assert.strictEqual(result.changed, false);
});

assert.strictEqual(
  banner.isCampaignActive(Date.parse("2026-09-17T23:59:59+03:00")),
  true
);
assert.strictEqual(
  banner.isCampaignActive(Date.parse("2026-09-18T00:00:00+03:00")),
  false
);

assert.strictEqual(banner.isCoordinatedMailbox("oguz.sahbazer@vito.com.tr"), true);
assert.strictEqual(banner.isCoordinatedMailbox("OGUZ.SAHBAZER@VITO.COM.TR"), true);
assert.strictEqual(banner.isCoordinatedMailbox("another.user@vito.com.tr"), false);

console.log("All ViTO banner runtime tests passed.");
