/**
 * autoTracker.js - Automatic email-to-deal tracking via Graph webhooks.
 *
 * When a user enables auto-tracking, Graph API sends webhook notifications
 * for every new email. This module processes those notifications and
 * automatically links emails to their associated HubSpot deals.
 *
 * Key design decision: only tracks threads that the user has explicitly
 * linked to a deal (via the "Link Email" action in Sphinx). No guessing
 * based on contact matching - this prevents false positives from linking
 * unrelated emails to the wrong deals.
 *
 * Safety: validates that at least one email participant is a contact on
 * the target deal before linking, because Graph API can group unrelated
 * emails under the same conversationId if subjects are similar.
 */
"use strict";

const tokenStore    = require("./tokenStore");
const subscriptions = require("./subscriptions");
const graph         = require("./graph");
const hubspot       = require("./hubspot");
const dedupStore    = require("./dedupStore");

// ── processNotification ──────────────────────────────────────
// Called from the webhook handler for each valid notification.
// Fetches the email, matches contacts to deals, and auto-links.
async function processNotification(notification) {
  const subscriptionId = notification.subscriptionId;
  const resourceId     = notification.resourceData?.id;

  if (!subscriptionId || !resourceId) {
    console.warn("[autoTracker] Notification missing subscriptionId or resourceData.id");
    return;
  }

  // Validate resourceId format — alphanumeric + safe chars only (same as isValidGraphId in index.js)
  if (typeof resourceId !== "string" || resourceId.length === 0 || resourceId.length > 500 || !/^[A-Za-z0-9\-_/+=.]+$/.test(resourceId)) {
    console.warn(`[autoTracker] Invalid resourceId format: ${String(resourceId).substring(0, 50)}`);
    return;
  }

  // 1. Identify which user this subscription belongs to
  const userId = subscriptions.getUserBySubscriptionId(subscriptionId);
  if (!userId) {
    console.warn(`[autoTracker] Unknown subscription ${subscriptionId} — ignoring`);
    return;
  }

  // 2. Get a valid access token for this user.
  // Prefer the cached token (valid ~1h), fall back to refresh token if expired.
  const user = tokenStore.getUser(userId);
  if (!user) {
    console.warn(`[autoTracker] No stored user data for ${userId} — skipping`);
    return;
  }

  let accessToken;
  // Try cached access token first (valid ~1h from when auto-track was enabled)
  if (user.accessToken && user.accessTokenExpiresAt && new Date(user.accessTokenExpiresAt).getTime() > Date.now()) {
    accessToken = user.accessToken;
  } else if (user.refreshToken) {
    // Fall back to refresh token
    try {
      const tokens = await graph.refreshAccessToken(user.refreshToken);
      accessToken = tokens.accessToken;
      // Cache new tokens - set expiry 5 min early (55 min) as safety margin
      tokenStore.saveUser(userId, {
        refreshToken: tokens.refreshToken,
        accessToken: tokens.accessToken,
        accessTokenExpiresAt: new Date(Date.now() + 55 * 60 * 1000).toISOString(),
      });
    } catch (err) {
      console.error(`[autoTracker] Token refresh failed for user ${userId}:`, err.message);
      return;
    }
  } else {
    console.warn(`[autoTracker] No valid token for user ${userId} — skipping`);
    return;
  }

  // 3. Fetch the full email
  let message;
  try {
    message = await graph.getMessageById(accessToken, resourceId);
  } catch (err) {
    console.warn(`[autoTracker] Failed to fetch message ${resourceId}:`, err.message);
    return;
  }

  if (!message) {
    // null = draft, skip silently
    return;
  }

  // 4. Dedup check
  const msgKey = message.internetMessageId || message.id;
  if (dedupStore.hasLinked(userId, msgKey)) {
    return; // already processed
  }

  // 5. Check if this email's conversation thread has been manually linked to a deal.
  //    Only conversations explicitly linked by the user (via "Link Email") are tracked.
  //    Before linking, verify participant overlap with deal contacts to prevent
  //    false matches from Graph's conversationId grouping (same subject, different people).
  let linkedAny = false;

  if (message.conversationId) {
    const dealId = dedupStore.getConversationDeal(message.conversationId);

    // Participant guard: check that at least one person on this email
    // (from, to, cc) is actually a contact on the deal.
    if (dealId) {
      try {
        const dealContacts = await hubspot.getDealContacts(dealId);
        const dealEmails = new Set(dealContacts.map(c => c.email.toLowerCase()));
        const emailParticipants = (message.allParticipants || []).map(e => e.toLowerCase());
        // Exclude the user themselves — they're on every email but shouldn't count
        const userEmailLower = (user.email || "").toLowerCase();
        const externalParticipants = emailParticipants.filter(e => e && e !== userEmailLower);
        const hasOverlap = externalParticipants.some(e => dealEmails.has(e));

        if (!hasOverlap && dealEmails.size > 0) {
          console.warn(`[autoTracker] Participant mismatch — skipping "${message.subject}". Deal contacts: [${[...dealEmails].join(", ")}], Email participants: [${externalParticipants.join(", ")}]`);
          dedupStore.markLinked(userId, msgKey);
          return;
        }
      } catch (err) {
        // If we can't fetch deal contacts, fall through and link anyway
        // (better to over-link than silently drop emails from a tracked thread)
        console.warn(`[autoTracker] Could not verify participants for deal ${dealId}:`, err.message);
      }
    }

    if (dealId && !dedupStore.hasDealLink(msgKey, dealId)) {
      try {
        await hubspot.linkEmailToDeal(
          {
            subject:   message.subject,
            bodyHtml:  message.bodyHtml,
            bodyText:  message.bodyText,
            fromEmail: message.fromEmail,
            toEmail:   message.toEmails,
            sentAt:    message.sentAt,
          },
          dealId,
          null,
          { userEmail: user.email }
        );

        linkedAny = true;
        dedupStore.markDealLink(msgKey, dealId);
        console.log(`[autoTracker] Linked email "${message.subject}" → deal ${dealId} via conversation mapping`);
      } catch (err) {
        console.error(`[autoTracker] linkEmailToDeal failed for deal ${dealId}:`, err.message);
      }
    }
  }

  // 6. Mark as processed (even if no mapping found — avoid re-checking)
  dedupStore.markLinked(userId, msgKey);

  if (linkedAny) {
    console.log(`[autoTracker] Successfully auto-linked message ${msgKey}`);
  }
}

module.exports = { processNotification };
