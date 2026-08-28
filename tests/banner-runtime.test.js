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

withActiveCampaign(function () {
  var wrongOrder = banner.buildBannerHtml() + "<div>Personal signature</div>";
  var result = banner.insertBannerIntoFullBody(wrongOrder);
  assert.strictEqual(result.changed, true);
  assert.ok(result.html.indexOf("Personal signature") < result.html.indexOf("data-vito-mail-banner"));
  assert.strictEqual(
    (result.html.match(/data-vito-mail-banner=/g) || []).length,
    1,
    "The current compose segment must contain one banner after repositioning."
  );
});

withActiveCampaign(function () {
  var wrongOrderWithQuotedBanner = [
    banner.buildBannerHtml(),
    "<div>Current personal signature</div>",
    '<div id="divRplyFwdMsg">',
    banner.buildBannerHtml(),
    "<div>Old message</div>",
    "</div>"
  ].join("");
  var result = banner.insertBannerIntoFullBody(wrongOrderWithQuotedBanner);
  assert.strictEqual(result.changed, true);
  assert.ok(result.html.indexOf("Current personal signature") < result.html.indexOf("data-vito-mail-banner"));
  assert.strictEqual(
    (result.html.match(/data-vito-mail-banner=/g) || []).length,
    2,
    "The current banner must move while the quoted banner remains untouched."
  );
});

assert.strictEqual(banner.composeHasMeaningfulContent("<html><body><p>&nbsp;</p></body></html>"), false);
assert.strictEqual(
  banner.composeHasMeaningfulContent("<html><body><div>Saygılarımla/Best Regards</div></body></html>"),
  true
);
assert.strictEqual(
  banner.composeHasMeaningfulContent("<html><body><img src=\"signature.png\"></body></html>"),
  true
);

assert.strictEqual(
  banner.isCampaignActive(Date.parse("2026-09-17T23:59:59+03:00")),
  true
);
assert.strictEqual(
  banner.isCampaignActive(Date.parse("2026-09-18T00:00:00+03:00")),
  false
);

console.log("All ViTO banner runtime tests passed.");
