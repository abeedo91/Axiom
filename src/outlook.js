import { PublicClientApplication, InteractionRequiredAuthError } from '@azure/msal-browser'

// Azure app registration values
const CLIENT_ID = 'd47432ed-b056-481f-b98c-676c693a117b'
const TENANT_ID = 'common' // 'common' allows personal + work accounts
const REDIRECT_URI = window.location.origin

const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: REDIRECT_URI,
  },
  cache: {
    cacheLocation: 'localStorage',
    storeAuthStateInCookie: false,
  },
}

const SCOPES = ['User.Read', 'Mail.Read', 'Mail.ReadWrite', 'offline_access']

let msalInstance = null
let msalReady = false

export async function initMsal() {
  if (msalReady) return msalInstance
  msalInstance = new PublicClientApplication(msalConfig)
  await msalInstance.initialize()
  // Handle redirect response (in case we return from login redirect)
  await msalInstance.handleRedirectPromise()
  msalReady = true
  return msalInstance
}

export async function signIn() {
  await initMsal()
  try {
    const result = await msalInstance.loginPopup({ scopes: SCOPES, prompt: 'select_account' })
    return result.account
  } catch (err) {
    if (err.message?.includes('popup')) {
      // Popup blocked - fall back to redirect
      await msalInstance.loginRedirect({ scopes: SCOPES })
      return null
    }
    throw err
  }
}

export async function signOut() {
  await initMsal()
  const accounts = msalInstance.getAllAccounts()
  if (accounts.length > 0) {
    await msalInstance.logoutPopup({ account: accounts[0] })
  }
}

export async function getAccount() {
  await initMsal()
  const accounts = msalInstance.getAllAccounts()
  return accounts.length > 0 ? accounts[0] : null
}

async function getToken() {
  await initMsal()
  const account = await getAccount()
  if (!account) throw new Error('Not signed in')

  try {
    const result = await msalInstance.acquireTokenSilent({ scopes: SCOPES, account })
    return result.accessToken
  } catch (err) {
    if (err instanceof InteractionRequiredAuthError) {
      const result = await msalInstance.acquireTokenPopup({ scopes: SCOPES, account })
      return result.accessToken
    }
    throw err
  }
}

// Fetch recent emails from inbox
export async function listEmails(top = 15) {
  const token = await getToken()
  const url = `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=${top}&$select=id,subject,from,bodyPreview,receivedDateTime,isRead,conversationId&$orderby=receivedDateTime desc`

  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  })

  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(err?.error?.message || `Failed to fetch emails (${res.status})`)
  }

  const data = await res.json()
  return data.value || []
}

// Fetch the full body of a single email
export async function getEmailBody(messageId) {
  const token = await getToken()
  const url = `https://graph.microsoft.com/v1.0/me/messages/${messageId}?$select=id,subject,from,toRecipients,body,receivedDateTime,conversationId`

  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  })

  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(err?.error?.message || `Failed to load email (${res.status})`)
  }

  return await res.json()
}

// Create a draft reply to a message
export async function createDraftReply(messageId, replyBody) {
  const token = await getToken()

  // Step 1: Create a reply draft
  const createRes = await fetch(
    `https://graph.microsoft.com/v1.0/me/messages/${messageId}/createReply`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
    }
  )

  if (!createRes.ok) {
    const err = await createRes.json().catch(() => ({}))
    throw new Error(err?.error?.message || `Failed to create draft (${createRes.status})`)
  }

  const draft = await createRes.json()

  // Step 2: Update the draft body with our AI-generated reply
  const updateRes = await fetch(
    `https://graph.microsoft.com/v1.0/me/messages/${draft.id}`,
    {
      method: 'PATCH',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        body: {
          contentType: 'HTML',
          content: replyBody.replace(/\n/g, '<br/>'),
        },
      }),
    }
  )

  if (!updateRes.ok) {
    const err = await updateRes.json().catch(() => ({}))
    throw new Error(err?.error?.message || `Failed to update draft (${updateRes.status})`)
  }

  return draft
}

// Strip HTML for Claude to read
export function stripHtml(html) {
  if (!html) return ''
  const tmp = document.createElement('div')
  tmp.innerHTML = html
  return (tmp.textContent || tmp.innerText || '').trim()
}
