'use strict';
const functions = require('firebase-functions');
const { WebhookClient } = require('dialogflow-fulfillment');
const http = require('http');
const querystring = require('querystring');
const url = require('url');
const WebSocket = require('ws');
const request = require('request');
const cheerio = require('cheerio');
const isProd = true;
const baseURI = isProd ? 'platform.quip.com' : 'platform.docker.qa';
const baseURL = isProd ? `https://${baseURI}` : `http://${baseURI}`;
const basePort = isProd ? 443 : 10000;
const httpLib = http;
const { Card, Suggestion } = require('dialogflow-fulfillment');
process.env.DEBUG = 'dialogflow:debug';
const YOUR_QUIP_ACCESS_TOKEN = process.env.YOUR_QUIP_ACCESS_TOKEN;

/**
 * A Quip API client.
 *
 * To make API calls that access Quip data, initialize with an accessToken.
 *
 *    const quip = require('quip');
 *    const client = quip.Client({accessToken: '...'});
 *    const user = await client.getAuthenticatedUser();
 *
 * To generate authorization URLs, i.e., to implement OAuth login, initialize
 * with a clientId and and clientSecret.
 *
 *    const quip = require('quip');
 *    const client = quip.Client({clientId: '...', clientSecret: '...'});
 *    response.writeHead(302, {
 *      'Location': client.getAuthorizationUrl()
 *    });
 *
 * @param {{accessToken: (string|undefined),
 *          clientId: (string|undefined),
 *          clientSecret: (string|undefined)}} options
 * @constructor
 */
function Client(options) {
  this.accessToken = options.accessToken;
  this.clientId = options.clientId;
  this.clientSecret = options.clientSecret;
}

/**
 * Returns the URL the user should be redirected to to sign in.
 *
 * @param {string} redirectUri
 * @param {string=} state
 */
Client.prototype.getAuthorizationUrl = function(redirectUri, state) {
  return (
    `${baseURL}:${basePort}/1/oauth/login?` +
    querystring.stringify({
      redirect_uri: redirectUri,
      state: state,
      response_type: 'code',
      client_id: this.clientId
    })
  );
};

/**
 * Exchanges a verification code for an access_token.
 *
 * Once the user is redirected back to your server from the URL
 * returned by `getAuthorizationUrl`, you can exchange the `code`
 * argument for an access token with this method.
 *
 * @param {string} redirectUri
 * @param {string} code
 * @return {Promise}
 */
Client.prototype.getAccessToken = function(redirectUri, code) {
  return this.call_(
    'oauth/access_token?' +
      querystring.stringify({
        redirect_uri: redirectUri,
        code: code,
        grant_type: 'authorization_code',
        client_id: this.clientId,
        client_secret: this.clientSecret
      })
  );
};

/**
 * @return {Promise}
 */
Client.prototype.getAuthenticatedUser = function() {
  return this.call_('users/current');
};

/**
 * @param {string} id
 * @return {Promise}
 */
Client.prototype.getUser = async function(id) {
  const users = await this.getUsers([id]);
  return users[id];
};

/**
 * @param {Array.<string>} ids
 * @return {Promise}
 */
Client.prototype.getUsers = function(ids) {
  return this.call_(
    'users/?' +
      querystring.stringify({
        ids: ids.join(',')
      })
  );
};

/**
 * @return {Promise}
 */
Client.prototype.getContacts = function() {
  return this.call_('users/contacts');
};

/**
 * @param {string} id
 * @return {Promise}
 */
Client.prototype.getFolder = async function(id) {
  const folders = await this.getFolders([id]);
  return folders[id];
};

/**
 * @param {Array.<string>} ids
 * @return {Promise}
 */
Client.prototype.getFolders = function(ids) {
  return this.call_(
    'folders/?' +
      querystring.stringify({
        ids: ids.join(',')
      })
  );
};

/**
 * @param {{title: string,
 *          parentId: (string|undefined),
 *          color: (Color|undefined),
 *          memberIds: (Array.<string>|undefined)}} options
 * @return {Promise}
 */
Client.prototype.newFolder = function(options) {
  const args = {
    title: options.title,
    parent_id: options.parentId,
    color: options.color
  };
  if (options.memberIds) {
    args['member_ids'] = options.memberIds.join(',');
  }
  return this.call_('folders/new', args);
};

/**
 * @param {{folderId: string,
 *          title: (string|undefined),
 *          color: (Color|undefined)}} options
 * @return {Promise}
 */
Client.prototype.updateFolder = function(options) {
  const args = {
    folder_id: options.folderId,
    title: options.title,
    color: options.color
  };
  return this.call_('folders/update', args);
};

/**
 * @param {{folderId: string,
 *          memberIds: Array.<string>}} options
 * @return {Promise}
 */
Client.prototype.addFolderMembers = function(options) {
  const args = {
    folder_id: options.folderId,
    member_ids: options.memberIds.join(',')
  };
  return this.call_('folders/add-members', args);
};

/**
 * @param {{folderId: string,
 *          memberIds: Array.<string>}} options
 * @return {Promise}
 */
Client.prototype.removeFolderMembers = function(options) {
  const args = {
    folder_id: options.folderId,
    member_ids: options.memberIds.join(',')
  };
  return this.call_('folders/remove-members', args);
};

/**
 * @param {{threadId: string,
 *          maxUpdatedUsec: (number|undefined),
 *          count: (number|undefined)}} options
 * @return {Promise}
 */
Client.prototype.getMessages = function(options) {
  return this.call_(
    'messages/' +
      options.threadId +
      '?' +
      querystring.stringify({
        max_updated_usec: options.maxUpdatedUsec,
        count: options.count
      })
  );
};

/**
 * @param {{threadId: string,
 *          frame: string,
 *          content: string,
 *          parts: Array<Array<string>>
 *          attachments: Array<string>
 *          silent: (boolean|undefined)
 *          annotationId: string,
 *          sectionId: string}} options
 * @return {Promise}
 */
Client.prototype.newMessage = function(options) {
  const args = {
    thread_id: options.threadId,
    frame: options.frame,
    content: options.content,
    parts: options.parts,
    attachments: options.attachments,
    silent: options.silent ? 1 : undefined,
    annotation_id: options.annotationId,
    section_id: options.sectionId
  };
  return this.call_('messages/new', args);
};

/**
 * @param {string} id
 * @return {Promise}
 */
Client.prototype.getThread = async function(id) {
  const threads = await this.getThreads([id]);
  return threads[id];
};

/**
 * @param {Array.<string>} ids
 * @return {Promise}
 */
Client.prototype.getThreads = function(ids) {
  return this.call_(
    'threads/?' +
      querystring.stringify({
        ids: ids.join(',')
      })
  );
};

/**
 * @param {{maxUpdatedUsec: (number|undefined),
 *          count: (number|undefined)}?} options
 * @return {Promise}
 */
Client.prototype.getRecentThreads = function(options) {
  return this.call_(
    'threads/recent?' +
      querystring.stringify(
        options
          ? {
              max_updated_usec: options.maxUpdatedUsec,
              count: options.count
            }
          : {}
      )
  );
};

/**
 * @param {{content: string,
 *          title: (string|undefined),
 *          format: (string|undefined),
 *          memberIds: (Array.<string>|undefined)}} options
 * @return {Promise}
 */
Client.prototype.newDocument = function(options) {
  const args = {
    content: options.content,
    title: options.title,
    format: options.format
  };
  if (options.memberIds) {
    args['member_ids'] = options.memberIds.join(',');
  }
  return this.call_('threads/new-document', args);
};

/**
 * @param {{query: string,
 *          count: (number|undefined),
 *          onlyMatchTitles: (boolean|undefined),
 * @return {Promise}
 */
Client.prototype.findDocument = function(options) {
  return this.call_(
    'threads/search?' +
      querystring.stringify(
        options
          ? {
              query: options.query,
              count: options.count,
              only_match_titles: options.onlyMatchTitles
            }
          : {}
      )
  );
};

/**
 * @param {{threadId: string,
 *          content: string,
 *          operation: (Operation|undefined),
 *          format: (string|undefined),
 *          sectionId: (string|undefined)}} options
 * @return {Promise}
 */
Client.prototype.editDocument = function(options) {
  const args = {
    thread_id: options.threadId,
    content: options.content,
    location: options.operation,
    format: options.format,
    section_id: options.sectionId
  };
  return this.call_('threads/edit-document', args);
};

/**
 * @param {{threadId: string,
 *          memberIds: Array.<string>}} options
 * @return {Promise}
 */
Client.prototype.addThreadMembers = function(options) {
  const args = {
    thread_id: options.threadId,
    member_ids: options.memberIds.join(',')
  };
  return this.call_('threads/add-members', args);
};

/**
 * @param {{threadId: string,
 *          memberIds: Array.<string>}} options
 * @return {Promise}
 */
Client.prototype.removeThreadMembers = function(options) {
  const args = {
    thread_id: options.threadId,
    member_ids: options.memberIds.join(',')
  };
  return this.call_('threads/remove-members', args);
};

/**
 * @param {{url: string,
 *          threadId: string}} options
 * @return {Promise}
 */
Client.prototype.addBlobFromURL = async function(options) {
  return this.call_(`blob/${options.threadId}`, {
    blob: request(options.url)
  });
};

/**
 * @param {{path: string,
 *          threadId: string}} options
 * @return {Promise}
 */
Client.prototype.addBlobFromPath = async function(options) {
  return this.call_(`blob/${options.threadId}`, {
    blob: fs.createReadStream(options.path)
  });
};

/**
 * @return {Promise<WebSocket>}
 */
Client.prototype.connectWebsocket = function() {
  return this.call_('websockets/new').then(newSocket => {
    if (!newSocket || !newSocket.url) {
      throw new Error(newSocket ? newSocket.error : 'Request failed');
    }
    const urlInfo = url.parse(newSocket.url);
    const ws = new WebSocket(newSocket.url, {
      origin: `${urlInfo.protocol}//${urlInfo.hostname}`
    });
    return ws;
  });
};

/**
 * @param {string} path
 * @param {Object.<string, *>=} postArguments
 * @return {Promise}
 */
Client.prototype.call_ = function(path, postArguments) {
  const requestOptions = {
    uri: `${baseURL}:${basePort}/1/${path}`,
    headers: {}
  };
  if (this.accessToken) {
    requestOptions.headers['Authorization'] = 'Bearer ' + this.accessToken;
  }
  if (postArguments) {
    const formData = {};
    for (let name in postArguments) {
      if (postArguments[name]) {
        formData[name] = postArguments[name];
      }
    }
    requestOptions.method = 'POST';
    requestOptions.formData = formData;
  } else {
    requestOptions.method = 'GET';
  }
  return new Promise((resolve, reject) => {
    const callback = (err, res, body) => {
      if (err) {
        return reject(err);
      }
      let responseObject = null;
      try {
        responseObject = /** @type {Object} */ (JSON.parse(body));
      } catch (err) {
        reject(`Invalid response for ${path}: ${body}`);
        return;
      }
      if (res.statusCode !== 200) {
        reject(new ClientError(res, responseObject));
      } else {
        resolve(responseObject);
      }
    };
    request(requestOptions, callback);
  });
};

/**
 * @param {http.IncomingMessage} httpResponse
 * @param {Object} info
 * @extends {Error}
 * @constructor
 */
function ClientError(httpResponse, info) {
  this.httpResponse = httpResponse;
  this.info = info;
}
ClientError.prototype = Object.create(Error.prototype);

const client = new Client({
  accessToken: YOUR_QUIP_ACCESS_TOKEN
});

function Quippy(options) {
  this.options = options;
}

Quippy.prototype.findDocumentByTitle = async function(title) {
  try {
    const foundResults = await client.findDocument({
      query: title,
      onlyMatchTitles: true
    });
    return foundResults;
  } catch (e) {
    console.log(e);
  }
};

Quippy.prototype.getDocumentContents = async function(id) {
  try {
    const foundThread = await client.getThread(id);
    const contents = foundThread.html.toString();
    return contents;
  } catch (e) {
    console.log(e);
  }
};

Quippy.prototype.editDocument = async function(id, documentContents, content) {
  try {
    const $ = cheerio.load(documentContents);
    const lastElementId = $('ul')
      .find('li')
      .last('span')
      .attr('id');
    const editedDocument = await client.editDocument({
      threadId: id,
      content: content,
      operation: 'AFTER_SECTION',
      sectionId: lastElementId
    });

    return editedDocument;
  } catch (e) {
    console.log(e);
  }
};

Quippy.prototype.findAndEditDocument = async function(title, contents) {
  const foundResults = await this.findDocumentByTitle(title);
  const numberOfResults = foundResults.length;
  if (numberOfResults > 0) {
    const documentId = foundResults[0].thread.id;
    const documentContents = await this.getDocumentContents(documentId);
    if (documentContents.length > 0) {
      const editedDocument = await this.editDocument(
        documentId,
        documentContents,
        contents
      );
      return editedDocument;
    }
  } else {
    return 'No Results Found';
  }
};

Quippy.prototype.getContentsByTitle = async function(title) {
  try {
    const doc = await this.findDocumentByTitle(title);
    const contents = await this.getDocumentContents(doc[0].thread.id);
    const $ = cheerio.load(contents);
    const textContents = $.text();
    if (textContents.length) {
      return textContents;
    } else {
      return 'Could not find that document';
    }
  } catch (e) {
    console.log(e);
  }
};

exports.dialogflowFirebaseFulfillment = functions.https.onRequest(
  (request, response) => {
    const agent = new WebhookClient({ request, response });
    const quippy = new Quippy();

    async function getQuipContentsHandler(agent) {
      const fileTitle = agent.parameters.title;
      const fileContents = await quippy.getContentsByTitle(fileTitle);
      agent.add(`${fileContents}`);
    }

    let intentMap = new Map();
    intentMap.set('getQuipContents', getQuipContentsHandler);
    agent.handleRequest(intentMap);
  }
);
