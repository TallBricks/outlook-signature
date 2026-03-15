// Tall Bricks - Outlook Signature Add-in
// Automatically inserts signature when compose window opens

// -----------------------------------------------
// CONFIG
// -----------------------------------------------
const LOGO_URL    = "https://www.tallbricks.com/wp-content/uploads/tb-email-logo.jpg";
const WEBSITE_URL = "https://www.tallbricks.com";
const ADDRESS1    = "Office 3105, Al Salam Tower, Al Sufouh Second,";
const ADDRESS2    = "Sheikh Zayed Road, Dubai, UAE";
// -----------------------------------------------

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        // Automatically run when add-in loads on compose
        insertSignature();
    }
});

async function insertSignature() {
    try {
        // Get access token to call Microsoft Graph
        const token = await getGraphToken();
        if (!token) {
            console.error("Could not obtain Graph token");
            insertFallbackSignature();
            return;
        }

        // Fetch user profile from Microsoft Graph
        const response = await fetch(
            "https://graph.microsoft.com/v1.0/me?$select=displayName,jobTitle,mobilePhone,businessPhones,mail",
            {
                headers: { Authorization: "Bearer " + token }
            }
        );

        if (!response.ok) {
            console.error("Graph API error:", response.status);
            insertFallbackSignature();
            return;
        }

        const user = await response.json();

        const name  = user.displayName  || "";
        const title = user.jobTitle     || "";
        const phone = user.mobilePhone  || (user.businessPhones && user.businessPhones[0]) || "";
        const email = user.mail         || "";

        const signatureHTML = buildSignatureHTML(name, title, phone, email);
        await setSignatureInCompose(signatureHTML);

    } catch (error) {
        console.error("Signature insert error:", error);
        insertFallbackSignature();
    }
}

function buildSignatureHTML(name, title, phone, email) {
    return `
<table cellpadding="0" cellspacing="0" border="0"
style="font-family: Arial, sans-serif; font-size: 13px; color: #333333; line-height: 1.5;">

  <tr>
    <td style="padding-bottom: 4px;">
      <strong style="font-size: 15px; color: #000000;">${name}</strong>
    </td>
  </tr>

  <tr>
    <td style="padding-bottom: 12px; color: #333333;">
      ${title}
    </td>
  </tr>

  <tr>
    <td style="padding-bottom: 16px;">
      <a href="mailto:${email}"
         style="color: #1a6fa3; text-decoration: underline;">${email}</a>
      &nbsp;&nbsp;${phone}
    </td>
  </tr>

  <tr>
    <td style="padding-bottom: 4px;">
      <img src="${LOGO_URL}" width="200" alt="Tall Bricks Real Estate LLC"
           style="display: block; border: 0;">
    </td>
  </tr>

  <tr>
    <td style="padding-bottom: 8px;">
      <hr style="border: none; border-top: 1px solid #cccccc; width: 200px; margin: 0; text-align: left;">
    </td>
  </tr>

  <tr>
    <td style="padding-bottom: 4px; font-size: 12px; color: #333333;">
      ${ADDRESS1}
    </td>
  </tr>

  <tr>
    <td style="padding-bottom: 12px; font-size: 12px; color: #333333;">
      ${ADDRESS2}
    </td>
  </tr>

  <tr>
    <td>
      <a href="${WEBSITE_URL}"
         style="color: #1a6fa3; text-decoration: underline; font-size: 13px;">www.tallbricks.com</a>
    </td>
  </tr>

</table>`;
}

async function setSignatureInCompose(html) {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.setSignatureAsync(
            html,
            { coercionType: Office.CoercionType.Html },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve();
                } else {
                    // setSignatureAsync may not be available on older builds
                    // Fall back to prependAsync
                    Office.context.mailbox.item.body.prependAsync(
                        "<br><br>" + html,
                        { coercionType: Office.CoercionType.Html },
                        (r) => {
                            if (r.status === Office.AsyncResultStatus.Succeeded) {
                                resolve();
                            } else {
                                reject(r.error);
                            }
                        }
                    );
                }
            }
        );
    });
}

async function getGraphToken() {
    return new Promise((resolve) => {
        Office.context.auth.getAccessTokenAsync(
            { allowSignInPrompt: true, allowConsentPrompt: true },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result.value);
                } else {
                    console.error("Token error:", result.error);
                    resolve(null);
                }
            }
        );
    });
}

function insertFallbackSignature() {
    // If Graph fails, insert a plain signature without personalization
    const fallback = buildSignatureHTML(
        "Your Name",
        "Your Title",
        "",
        ""
    );
    setSignatureInCompose(fallback).catch(console.error);
}
