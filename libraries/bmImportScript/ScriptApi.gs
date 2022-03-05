

/**
 * @typdef {GotReponse} 
 * @property {boolean} success whether it worked
 * @property {object|string} data the parsed data if it was an object or bare data if not
 * @property {number} code the http status code
 * @property {Error} extended error info
 * @property {boolean} parsed whether it was parsed successfilly
 * @property {object} headers the http headers
 * @property {string} content the unparsed content
 * @property {boolean} cached whether it came from cache
 * @property {function} throw throw an error if failed or return the result - can used like getProject(params).throw()
 */



class ScriptApi {
  /**
   * needs 2 scopes
   * https://www.googleapis.com/auth/script.external_request
   * https://www.googleapis.com/auth/script.projects
   * @param {object} param
   * @param {function} param.tokenService probably Script_App .getOAuthToken, with above scopes
   * @param {function} param.fetcher probably Url_Fetch_App .fetch to fetch externally
   * @param {CacheStore} [param.cacheStore=null] probably something like Cache_Service.getUserCache() if null/no caching
   * @param {number} [param.cacheSeconds=360] how to long to live in cache
   */
  constructor({ tokenService, fetcher, cacheStore = null, cacheSeconds = 60 * 60 }) {
    if (!fetcher) throw new Error('there must be a fetcher function')
    if (!tokenService) throw new Error('there must be a tokenService function')
    this.fetcher = new bmCrusher.Fetcher({
      tokenService,
      fetcher
    })
    this.cacher = new Cacher({ cacheStore, cacheSeconds })
    const _urlCleaner = (url) => url.replace(/\/+/g, '/').replace(/\/$/, '')
    const _baseUrl = "script.googleapis.com/v1"
    const joinParams = (params) => params && params.length ? ('?' + params.map(f=>{
      return `${f[0]}=${encodeURI(f[1].toString())}`
    }).join('&')) : ''
    this.makeUrl = (tail = '', params) => 'https://' + _urlCleaner(_baseUrl + '/' + tail) + joinParams(params)
    this.cacheSeconds = cacheSeconds
  }


  get fileTypes() {
    return [
      'SERVER_JS',
      'HTML',
      'JSON'
    ]
  }

  get collisions() {
    return [
      'abort',
      'replace',
      'skip',
      'rename'
    ]
  }
  /**
   * create a project
   * @param {object} param
   * @param {string} param.name name of script
   * @param {string} [param.title] container id 
   * @return {GotResponse}
   */
  createProject({ title, parentId }) {
    return this.got({
      url: this.makeUrl(`projects`),
      options: {
        method: 'POST',
        payload: {
          title,
          parentId
        }
      }
    })
  }

  /**
   * check the file makes sense
   */
  checkFileType(fileType) {
    if (this.fileTypes.indexOf(fileType) === -1) throw new Error(`Filetype ${fileType} not one of ${this.fileTypes}`)
  }
  /**
 * check the collission makes sense
 */
  checkFileCollision(collision) {
    if (this.collisions.indexOf(collision) === -1) throw new Error(`Collision ${collision} not one of ${this.collisions}`)
  }

  sameFile (a, b)  {
    return a.name === b.name && a.type === b.type
  }
  sameContent (a,b) {
    return this.sameFile(a,b) && a.source === b.source
  }
  /**
   * add files to a project
   * @param {object} param
   * @param {string} param.scriptId id of script
   * @param {File[]} [param.files=[]] array of files to add
   * @param {string} [param.collision = 'abort'] strategy for collision
   * @param {boolean} [param.clear = false] whether to clear out project before adding files
   * @param {boolean} [param.keepManifest = false] whether to keep the existing manifest (if there is one)
   * @return {GotResponse} 
   */
  addFiles({ scriptId, files =[], collision = 'abort', clear = false, keepManifest = false }) {
    // make sure types are good
    this.checkFileCollision(collision)
    files.every(f => this.checkFileType(f.type))

    // get current state of project
    const current = this.getProjectContent({ scriptId, noCache: true })
    if (!current.success) return current


    let currentFiles = current.data.files;
    // special handling of manifest file
    const currentManifest = currentFiles.find(f=>this.isManifest(f))
    const newManifest = files.find(f=>this.isManifest(f))
    const manifest = keepManifest ? (currentManifest || newManifest) : (newManifest || currentManifest)
    currentFiles = currentFiles.filter(f=>!this.isManifest(f))
    files = files.filter(f=>!this.isManifest(f))

    if (clear) {
      currentFiles = files
    } else {

      // if there's no change in content, we'll ignore completely
      files = files.filter(f=>!currentFiles.some(c=>this.sameContent(f,c)))

      switch (collision) {
        case 'abort':
          const collisions = currentFiles.filter(c => files.find(f => this.sameFile(f, c)))
          if (collisions.length) {
            return {
              ...current,
              success: false,
              collision,
              extended: new Error(`File collision(s) detected ${collisions.map(f=>f.name+"("+f.type+")").join(",")} : 'use collision: replace/skip/rename`)
            }
          }
          Array.prototype.push.apply(currentFiles, files)
          break;

        case 'replace':
          currentFiles = currentFiles.filter(c => !files.find(f => this.sameFile(f, c))).concat(files)
          break;

        case 'skip':
          currentFiles = currentFiles.concat(files.filter(c => !currentFiles.find(f => this.sameFile(f, c))))
          break;

        case 'rename':
          currentFiles = files.reduce((p, c) => {
            let { name } = c
            let index = 0
            while (p.find(f => this.sameFile(f, { ...c, name }))) {
              name = c.name + '_' + index++
            }
            p.push({
              ...c,
              name
            })
            return p
          }, currentFiles)
          break;

        default:
          throw new Error('damn')
      }
    }
    return this.updateProjectContent({ scriptId, files: currentFiles.concat([manifest]) })
  }

  /**
   * just chek that a file is a manifest file
   */
  isManifest ({name, type}) {
    return name === 'appsscript' && type === "JSON"
  } 

  isServer ({type}) {
    return type === "SERVER_JS"
  } 

  isClient ({type}) {
    return type === "HTML"
  } 

  /**
   * get the project deployments
   * @param {object} params
   * @param {string} params.scriptid the scriptid
   * @param {boolean} params.noCache whether to skip getting from cache
   * @param {number} params.cacheSeconds no of seconds to write to cache for after api get
   * @returns {GotResponse}
   */
  listProjectDeployments ({ scriptId, noCache = false, cacheSeconds }) {

    let pageToken = null
    let allDeployments = null
    let pargs= [['pageSize', 100]]
    do {
      if (pageToken) Array.prototype.push.apply(pargs, [['pageToken', pageToken]])
      const result = this.got({
        url: this.makeUrl(`projects/${scriptId}/deployments`, pargs),
        noCache,
        cacheSeconds
      })
      // allow caller to throw
      if (!result.success) return result
      // pile them up
      const {data} = result
      const {deployments} = data
      pageToken = data.nextPageToken
      if(allDeployments) {
        Array.prototype.push.apply(allDeployments.data.deployments, deployments) 
      } else {
        allDeployments = result
      }
 
    } while (pageToken)
    return allDeployments
  }

  /**
   * get the project summary
   * @param {object} params
   * @param {string} params.scriptid the scriptid
   * @param {boolean} params.noCache whether to skip getting from cache
   * @param {number} params.cacheSeconds no of seconds to write to cache for after api get
   * @returns {GotResponse}
   */
  getProject({ scriptId, noCache = false, cacheSeconds }) {
    return this.got({
      url: this.makeUrl(`projects/${scriptId}`),
      noCache,
      cacheSeconds
    })
  }

  /**
   * get the project content
   * @param {object} params
   * @param {string} params.scriptid the scriptid
   * @param {boolean} [params.noCache=false] whether to skip getting from cache
   * @param {number} [params.cacheSeconds] no of seconds to write to cache for after api get
   * @param {boolean} [params.skipManifest=false] whether to get the manifest too
   * @param {string|number} [params.versionNumber=null] version number - null for latest
   * @returns {GotResponse}
   */
  getProjectContent({ scriptId, noCache = false, cacheSeconds, skipManifest = false, versionNumber = null }) {
    const result =  this.got({
      url: this.makeUrl(`projects/${scriptId}/content`, versionNumber ? [['versionNumber', versionNumber]] : null),
      noCache,
      cacheSeconds
    })
    // sometimes we really don't want the manifest
    if (result.success && skipManifest) result.data.files = result.data.files.filter(f=>!this.isManifest(f))
    return result
  }

  /**
   * update the project content
   * @param {object} params
   * @param {string} params.scriptid the scriptid
   * @param {File[]} params.files the file content to update
   * @returns {GotResponse}
   */
  updateProjectContent({ scriptId, files }) {

    // need to strip off junk from previous fetch
    const payload = {
      files: files.map(f => ({
        name: f.name,
        type: f.type,
        source: f.source
      }))
    }

    const result = this.got({
      url: this.makeUrl(`projects/${scriptId}/content`),
      options: {
        method: 'PUT',
        payload
      }
    })

    return result;
  }

  /**
  * execute a urlfetch
  * @param {string} url the url
  * @param {object} options any additional options
  * @param {boolean} noCache whether to skip getting from cache
  * @param {number} cacheSeconds no of seconds to write to cache for after api get
  * @return {GotReponse} a standard response
  */
  got({ url, options, noCache = false, cacheSeconds }) {
    cacheSeconds = cacheSeconds || this.cacheSeconds
    options = {
      method: "GET",
      ...options
    }
    // use this key for caching
    const key = url 

    const method = options.method.toLowerCase()
    const getting = method === 'get'

    // we'll scrap the cache if it's any kind of write
    if (!getting) {
      this.cacher.remove(key)
    }

    // sort out json payload
    const posting = options.payload && typeof options.payload === 'object' && ['post', 'put'].indexOf(method) !== -1
    if (posting) {
      options.contentType = 'application/json'
      options.payload = JSON.stringify(options.payload)
    }

    // see if its in cache if we're getting
    const cached = getting && !noCache && this.cacher.get(key)

    // wasnt in cache
    const result = cached || this.fetcher.got(url, options)
    result.cached = Boolean(cached)
    
    // add a throw method for errors
    result.throw = () => {
      // just return the result as is if all is good
      if (result.success) return result

      // which error to throw
      if (result.data && result.data.error && result.data.error.message) {
        throw new Error(result.data.error.message)
      } else {
        throw new Error(result.extended)
      }
    }
    // write away to cache for next time - if an error, we don't write (already know there's nothing in cache so no need to clear)
    if (getting && result.success) {
      this.cacher.put(key, result, cacheSeconds)
    }
    return result
  }

}

/** 
 * export new instance
 * needs thee scopes
 * https://www.googleapis.com/auth/script.external_request
 * https://www.googleapis.com/auth/script.projects
 * @param {object} param
 * @param {function} param.tokenProvider probably Script_App .getOAuthToken, with above scopes
 * @param {function} param.fetcher probably Url_Fetch_App .fetch to fetch externally
 * @param {CacheStore} [param.cacheStore=null] probably something like Cache_Service.getUserCache() if null/no caching
 */
var newScriptApi = (params) => new ScriptApi(params)