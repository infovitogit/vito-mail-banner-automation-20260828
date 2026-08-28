/*
 * ViTO Office 365 Mail Banner Automation
 * Independent from the ViTO signature pilot.
 * Preserves the current client signature and inserts the active campaign
 * immediately before the quoted reply/forward history.
 */
(function (root) {
  "use strict";

  var CAMPAIGN = {
    id: "vito-gastech-bangkok-2026",
    imageUrl: "https://infovitogit.github.io/vito-mail-banner-automation-20260828/assets/gastechbangkok2026banner.png",
    altText: "ViTO is attending GASTECH Bangkok 2026",
    displayWidth: 600,
    expiresAt: "2026-09-18T00:00:00+03:00"
  };

  var ORDERING = {
    signatureWaitMs: 6000,
    pollIntervalMs: 250
  };

  function log(level, message) {
    try {
      if (root.console && typeof root.console[level] === "function") {
        root.console[level]("VITO_BANNER|" + message);
      }
    } catch (_) {
      // Logging must never interrupt message compose.
    }
  }

  function isCampaignActive(nowMs) {
    var current = typeof nowMs === "number" ? nowMs : Date.now();
    return current < Date.parse(CAMPAIGN.expiresAt);
  }

  function buildBannerHtml() {
    return [
      '<div data-vito-mail-banner="', CAMPAIGN.id, '" ',
      'style="margin:12px 0 0 0;padding:0;line-height:0;">',
      '<img src="', CAMPAIGN.imageUrl, '" ',
      'width="', String(CAMPAIGN.displayWidth), '" ',
      'alt="', CAMPAIGN.altText, '" ',
      'style="display:block;width:', String(CAMPAIGN.displayWidth),
      'px;max-width:100%;height:auto;border:0;outline:none;text-decoration:none;" />',
      '</div>'
    ].join("");
  }

  function findReplyBoundary(html) {
    var source = String(html || "");
    var patterns = [
      /<a\b[^>]*\bname\s*=\s*["']?_MailOriginal["']?[^>]*>/i,
      /<div\b[^>]*\bid\s*=\s*["'](?:x_)?divRplyFwdMsg["'][^>]*>/i,
      /<div\b[^>]*\bclass\s*=\s*["'][^"']*\bms-outlook-mobile-reference-message\b[^"']*["'][^>]*>/i
    ];
    var boundary = -1;

    patterns.forEach(function (pattern) {
      var match = pattern.exec(source);
      if (match && (boundary === -1 || match.index < boundary)) {
        boundary = match.index;
      }
    });

    return boundary;
  }

  function findClosingBody(html) {
    var source = String(html || "");
    var match = /<\/body\s*>/i.exec(source);
    return match ? match.index : source.length;
  }

  function currentComposeSegment(html) {
    var source = String(html || "");
    var boundary = findReplyBoundary(source);
    return boundary >= 0 ? source.slice(0, boundary) : source;
  }

  function hasCurrentBanner(html) {
    return currentComposeSegment(html).indexOf('data-vito-mail-banner="' + CAMPAIGN.id + '"') >= 0;
  }

  function currentBannerPattern() {
    return new RegExp(
      '<div\\b(?=[^>]*\\bdata-vito-mail-banner\\s*=\\s*["\\\']' +
        CAMPAIGN.id +
        '["\\\'])[^>]*>[\\s\\S]*?<\\/div\\s*>',
      "gi"
    );
  }

  function removeCurrentBanner(html) {
    var source = String(html || "");
    var boundary = findReplyBoundary(source);
    var current = boundary >= 0 ? source.slice(0, boundary) : source;
    var quoted = boundary >= 0 ? source.slice(boundary) : "";
    return current.replace(currentBannerPattern(), "") + quoted;
  }

  function composeHasMeaningfulContent(html) {
    var source = stripDocumentWrapper(removeCurrentBanner(html));

    if (/<(?:img|table)\b/i.test(source)) {
      return true;
    }

    var text = source
      .replace(/<style\b[^>]*>[\s\S]*?<\/style\s*>/gi, "")
      .replace(/<[^>]+>/g, "")
      .replace(/&nbsp;|&#160;|&#xA0;/gi, "")
      .replace(/[\s\u200B-\u200D\uFEFF]+/g, "");

    return text.length > 0;
  }

  function insertBannerIntoFullBody(html) {
    var source = String(html || "");
    if (!isCampaignActive()) {
      return { changed: false, html: source };
    }

    var withoutCurrentBanner = removeCurrentBanner(source);
    var boundary = findReplyBoundary(withoutCurrentBanner);
    var insertAt = boundary >= 0 ? boundary : findClosingBody(withoutCurrentBanner);
    var updated =
      withoutCurrentBanner.slice(0, insertAt) +
      buildBannerHtml() +
      withoutCurrentBanner.slice(insertAt);
    return { changed: updated !== source, html: updated };
  }

  function stripDocumentWrapper(html) {
    var source = currentComposeSegment(html);
    source = source.replace(/^[\s\S]*?<body\b[^>]*>/i, "");
    source = source.replace(/<\/body\s*>[\s\S]*$/i, "");
    source = source.replace(/<\/html\s*>[\s\S]*$/i, "");
    return source;
  }

  function isMobileOutlook() {
    try {
      var platform = String(Office.context.platform || "").toLowerCase();
      return platform.indexOf("ios") >= 0 || platform.indexOf("android") >= 0;
    } catch (_) {
      return false;
    }
  }

  function getBodyHtml(item, currentReplyOnly) {
    return new Promise(function (resolve, reject) {
      var callback = function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(String(result.value || ""));
        } else {
          reject(new Error(result.error && result.error.message ? result.error.message : "body.getAsync failed"));
        }
      };

      if (
        currentReplyOnly &&
        Office.MailboxEnums &&
        Office.MailboxEnums.BodyMode &&
        typeof Office.MailboxEnums.BodyMode.HostConfig !== "undefined"
      ) {
        item.body.getAsync(
          Office.CoercionType.Html,
          { bodyMode: Office.MailboxEnums.BodyMode.HostConfig },
          callback
        );
      } else {
        item.body.getAsync(Office.CoercionType.Html, callback);
      }
    });
  }

  function waitForComposeContent(item, currentReplyOnly) {
    var startedAt = Date.now();

    return new Promise(function (resolve, reject) {
      function poll() {
        getBodyHtml(item, currentReplyOnly)
          .then(function (html) {
            if (composeHasMeaningfulContent(html)) {
              resolve({ html: html, ready: true });
              return;
            }

            if (Date.now() - startedAt >= ORDERING.signatureWaitMs) {
              resolve({ html: html, ready: false });
              return;
            }

            root.setTimeout(poll, ORDERING.pollIntervalMs);
          })
          .catch(reject);
      }

      poll();
    });
  }

  function setFullBodyHtml(item, html) {
    return new Promise(function (resolve, reject) {
      item.body.setAsync(
        html,
        { coercionType: Office.CoercionType.Html },
        function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(new Error(result.error && result.error.message ? result.error.message : "body.setAsync failed"));
          }
        }
      );
    });
  }

  function setMobileSignatureAndBanner(item, currentReplyHtml) {
    return new Promise(function (resolve, reject) {
      var currentSignature = stripDocumentWrapper(currentReplyHtml);
      if (currentSignature.indexOf('data-vito-mail-banner="' + CAMPAIGN.id + '"') >= 0) {
        resolve(false);
        return;
      }

      item.body.setSignatureAsync(
        currentSignature + buildBannerHtml(),
        { coercionType: Office.CoercionType.Html },
        function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(true);
          } else {
            reject(new Error(result.error && result.error.message ? result.error.message : "setSignatureAsync failed"));
          }
        }
      );
    });
  }

  function completeOnce(event) {
    var completed = false;
    return function () {
      if (completed) {
        return;
      }
      completed = true;
      try {
        event.completed();
      } catch (_) {
        // Outlook may already have closed the compose event.
      }
    };
  }

  function insertVitoMailBanner(event) {
    var done = completeOnce(event);

    if (!isCampaignActive()) {
      log("info", "CAMPAIGN_EXPIRED");
      done();
      return;
    }

    try {
      var item = Office.context.mailbox.item;
      var mobile = isMobileOutlook();

      waitForComposeContent(item, mobile)
        .then(function (bodyState) {
          var html = bodyState.html;
          log("log", bodyState.ready ? "SIGNATURE_READY" : "SIGNATURE_WAIT_TIMEOUT");

          if (mobile) {
            return setMobileSignatureAndBanner(item, html).then(function (changed) {
              log("log", changed ? "MOBILE_INSERT_OK" : "MOBILE_ALREADY_PRESENT");
            });
          }

          var result = insertBannerIntoFullBody(html);
          if (!result.changed) {
            log("log", "DESKTOP_ALREADY_PRESENT_OR_EXPIRED");
            return null;
          }

          return setFullBodyHtml(item, result.html).then(function () {
            log("log", "DESKTOP_INSERT_OK");
          });
        })
        .catch(function (error) {
          log("error", "INSERT_FAILED|" + (error && error.message ? error.message : String(error)));
        })
        .then(done, done);
    } catch (error) {
      log("error", "UNEXPECTED|" + (error && error.message ? error.message : String(error)));
      done();
    }
  }

  var api = {
    CAMPAIGN: CAMPAIGN,
    ORDERING: ORDERING,
    buildBannerHtml: buildBannerHtml,
    composeHasMeaningfulContent: composeHasMeaningfulContent,
    currentComposeSegment: currentComposeSegment,
    findReplyBoundary: findReplyBoundary,
    hasCurrentBanner: hasCurrentBanner,
    insertBannerIntoFullBody: insertBannerIntoFullBody,
    isCampaignActive: isCampaignActive,
    removeCurrentBanner: removeCurrentBanner,
    stripDocumentWrapper: stripDocumentWrapper
  };

  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }

  root.VitoBannerAutomation = api;

  if (typeof Office !== "undefined") {
    try {
      Office.actions.associate("insertVitoMailBanner", insertVitoMailBanner);
    } catch (_) {
      Office.onReady(function () {
        Office.actions.associate("insertVitoMailBanner", insertVitoMailBanner);
      });
    }
  }
})(typeof globalThis !== "undefined" ? globalThis : this);
