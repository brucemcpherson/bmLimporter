
/**
 * @typdef ParsedFile the compiled source
 * @property {File} file
 * @property {Ast} ast
 */
class Limporter {

  /*
   * needs 3 scopes
   * https://www.googleapis.com/auth/script.external_request
   * https://www.googleapis.com/auth/script.projects
   * https://www.googleapis.com/auth/script.deployments.readonly
   * constructor arguments are the same as bmImportScript
   * @param {object} params
   * @param {function} params.tokenService probably Script_App .getOAuthToken, with above scopes
   * @param {function} params.fetcher probably Url_Fetch_App .fetch to fetch externally
   * @param {CacheStore} [params.cacheStore=null] probably something like Cache_Service.getUserCache() if null/no caching
   * @param {number} [params.cacheSeconds=360] how to long to live in cache
   */
  constructor(params) {
    this.sapi = bmImportScript.newScriptApi(params)
    this.startedAt = new Date()
    this.versionTreatments = [
      { name: 'respect', description: 'respect the version numbers in the manifest' },
      { name: 'upgrade', description: 'upgrade to the latest deployed versions' },
      { name: 'head', description: 'use the latest version whether or not deployed' },
      { name: 'sloppyupgrade', description: 'use the latest deployed version if there is one - otherwise use the head' }
    ]
    // detect loopbacks with a credible depth
    this.MAX_DEPTH = 10
  }

  validateTreatments({ versionTreatment = 'respect' }) {
    const treatment = this.versionTreatments.find(f => f.name === versionTreatment)
    if (!treatment) throw new Error(`Version treatment ${versionTreatment} should be one of ${JSON.stringify(this.versionTreatments)}`)
    return treatment
  }
  /**
   * manifest structure
   * some from master .. ignore others
   * some decorte master
   * {
  "addOns": {
    object (AddOns)  
  },
  "dependencies": {
    object (Dependencies)
  },
  "exceptionLogging": string,
  "executionApi": {
    object (ExecutionApi)
  },
  "oauthScopes": [
    string
  ],
  "runtimeVersion": string,
  "sheets": {
    object (Sheets)
  },
  "timeZone": string,
  "urlFetchWhitelist": [
    string
  ],
  "webapp": {
    object (Webapp)
  }
}
*/

  mergeManifests(params, inlined) {

    // find the master manifest
    const master = {
      ...inlined.filter(f => !f.namespace.depth)[0].namespace.manifest
    }
    const original = JSON.parse(JSON.stringify(master))

    // if you don't upgrade to v8, there may be unexpected errors
    const { runtimeVersion = "V8" } = params;
    master.runtimeVersion = runtimeVersion

    // get all the others
    // &remove properties not inherited by the master manifest
    const others = inlined.filter(f => f.namespace.depth)
      .map(f => f.namespace.manifest)
      .map(f => ({
        dependencies: f.dependencies,
        oauthScopes: f.oauthScopes,
        sheets: f.sheets,
        urlFetchWhitelist: f.urlFetchWhitelist
      }))

    // decorate the master
    const mergedManifest = others.reduce((p, c) => {

      // dependencies are libraries and EnabledAdvancedService
      // we can ignore libraries as they'll be removed anyway
      if (c.dependencies && c.dependencies.enabledAdvancedServices && c.dependencies.enabledAdvancedServices.length) {
        p.dependencies = p.dependencies || {}
        p.dependencies.enabledAdvancedServices = p.dependencies.enabledAdvancedServices || []
        // this might cause a problem if multple versions are enabled with the same usersymbol 
        c.dependencies.enabledAdvancedServices.forEach(s => {
          const eas = p.dependencies.enabledAdvancedServices
            .find(({ serviceId, userSymbol, versionId }) => serviceId === s.serviceId && userSymbol === s.userSymbol && s.versionId === s.versionId)
          if (!eas) {
            // could be an ambigous user symbol
            const eus = p.dependencies.enabledAdvancedServices
              .find(({ userSymbol }) => serviceId === userSymbol === s.userSymbol)
            if (eus) {
              console.log(`skipping potentially ambiguous Advanced service ${s.userSymbol} in chid libraries`)
            } else {
              // good to add this one
              p.dependencies.enabledAdvancedServices.push(s)
            }
          }
        })
      }
      // oauth scopes are just added and dedupped
      p.oauthScopes = (p.oauthScopes || []).concat(c.oauthScopes || []).filter((f, i, a) => a.indexOf(f) === i)

      // sheets are all about macros - I guess there could be some in libraries
      if (c.sheets && c.sheets.macros && c.sheets.macros.length) {
        p.sheets = p.sheets || {}
        p.sheets.macros = p.sheets.macros || []
        c.sheets.macros.forEach(s => {
          const ks = p.sheets.macros.find(f => s.defaultShortcut)
          if (ks) {
            if (ks.menuName !== s.menuName) {
              console.log(`skipping potentially ambiguous macro service ${s.menuName} in chid libraries`)
            }
          } else {
            p.sheets.macros.push(s)
          }
        })
      }

      // url whitelists are just added and dedupped
      p.urlFetchWhitelist = (p.urlFetchWhitelist || []).concat(c.urlFetchWhitelist || []).filter((f, i, a) => a.indexOf(f) === i)

      return p
    }, master)

    // get rid of the libraries and unneeded references
    if (mergedManifest.dependencies && mergedManifest.dependencies.libraries) {
      delete mergedManifest.dependencies.libraries
    }
    if (mergedManifest.dependencies && !Object.keys(mergedManifest.dependencies).length) {
      delete mergedManifest.dependencies
    }
    if (mergedManifest.urlFetchWhitelist && !mergedManifest.urlFetchWhitelist.length) {
      delete mergedManifest.urlFetchWhitelist
    }

    return {
      mergedManifest,
      original
    }
  }


  getImportFolder(params) {
    return params.importFolder || "_bmlimport/"
  }

  getImportManifestName(params) {
    return this.makeImportName(params, '__manifest')
  }

  getGlobalsName(params) {
    return this.makeImportName(params, '__globals')
  }

  makeImportName(params, name) {
    return `${this.getImportFolder(params)}${name}`
  }

  makeGetters(names) {
    const { ns, s } = this.indenter
    const code = `var ${this.getterObject} = {${ns}` +
      names.map(f => `get ${f} () { ${ns}${s}return new ${f} () ${ns}}`).join("," + ns) +
      "\n}"
    return code
  }

  /**
   * revert to library version
   */
  revert(params) {

    // throw an error on failure as all should be present
    const folder = this.getImportFolder(params)
    const manifest = this.getImportManifestName(params)

    const project = this.sapi.getProject(params).throw()
    console.log(`reverting ${project.data.title} - ${project.data.scriptId}`)
    const content = this.sapi.getProjectContent(params).throw()

    const { files } = content.data

    // find any imported files except the original manifest
    const originalManifests = files.filter(f => f.name === manifest)
    if (originalManifests.length !== 1) {
      throw new Error(`Can't revert: Missing or multiple ${manifest}`)
    }
    const originalManifest = originalManifests[0]

    // get rid of everything imported and just keep the project files and original manifest
    const newFiles = files
      .filter(f => !(f.name.startsWith(folder) || this.sapi.isManifest(f)) || f.name === manifest)

    //rename the original to the new
    originalManifest.name = 'appsscript'
    if (!this.sapi.isManifest(originalManifest)) {
      throw new Error(`${JSON.stringify(originalManifest.name)} is invalid`)
    }
    // this is the reverted project
    return newFiles

  }

  getInlineProjectFiles(params) {

    // inline all the libraries
    const { inlined, scripts } = this.makeLibraries(params)

    // get all the manifests and merge them
    const merges = this.mergeManifests(params, inlined)

    // move all the namespaces to a file
    const files = inlined.reduce((p, { code, namespace }) => {
      if (namespace.depth) {
        p.push({
          name: this.makeImportName(params, namespace.namespaceName),
          type: "SERVER_JS",
          source: code
        })
      }
      else {
        if (namespace.children.length) {
          p.push({
            name: this.getGlobalsName(params),
            type: "SERVER_JS",
            source: this.makeInstanceOfLibraries(namespace.children)
          })
        }
        Array.prototype.push.apply(p, namespace.files)
      }
      return p
    }, [{
      name: this.makeImportName(params, this.getterObject),
      type: "SERVER_JS",
      source: this.makeGetters(Array.from(scripts.values()).map(f => f.namespaceName))
    }, {
      name: this.getImportManifestName(params),
      type: "JSON",
      source: JSON.stringify(merges.original, null, 2),
    }, {
      name: 'appsscript',
      source: JSON.stringify(merges.mergedManifest, null, 2),
      type: "JSON"
    }]).filter((f, i, a) => a.findIndex(g => this.sapi.sameContent(g, f)) === i)

    return files

  }

  /**
   * get just the manifest
   * @param {object} param 
   * @param {File[]} param.files the project content files
   * returns {File}
   */
  getManifest({ files }) {
    return files.find(f => this.sapi.isManifest(f))
  }

  /**
  * get just the serverside files
  * @param {object} param 
  * @param {File[]} param.files the project content files
  * returns {File[]}
  */
  getServers({ files }) {
    return files.filter(f => this.sapi.isServer(f))
  }

  /**
   * get just the clientside files
   * @param {object} param 
   * @param {File[]} param.files the project content files
   * returns {File[]}
   */
  getClients({ files }) {
    return files.filter(f => this.sapi.isClient(f))
  }

  /**
   * compile a source file
   * @param {File} file the file to be parsed
   * @return {ParsedFile} {ast, file}
   */
  getAst(file) {
    return {
      ast: bmAcorn.acorn.parse(file.source, {
        sourceType: "script",
        ecmaVersion: 12,
        allowReserved: true
      }), // parseScript
      file
    }
  }

  /**
   * document on source of file
   * @param {object} params
   * @param {File} params.file
   * @param {string} scriptId the script it came from
   * @param {number} [indent=2] how many to indent by
   * @returns {string} the commented string
   */
  commenter({ file }, indent = 2) {
    return `//--script file:${file.name}\n${this.spaces(indent) +
      file.source.replace(/\n/g, '\n' + this.spaces(indent))}\n//--end:${file.name}\n`
  }

  /**
   * document on source of file
   * @param {object} params
   * @param {string} params.title the project title
   * @param {string} params.scriptId the script it came from
   * @param {string} params.content the namespace content
   * @returns {string} the commented string
   */
  namespaceCommenter({ scriptId, title, versionNumber = 0, content }) {
    return `//--project:${scriptId} (${title}) version:${versionNumber || 'latest'}\n` +
      `//  imported by bmImportLibraries at ${this.startedAt.toUTCString()}\n${content}\n//--end project:${title}`
  }

  spaces(indent) {
    return ' '.repeat(indent)
  }

  /**
   * get interesting body
   * @param {ParsedFile} the file
   * @param {string} the type 'FunctionDeclaration'|'VariableDeclaration
   * @returns {object[]} the wualifying bodies
   */
  getBody(ast, type) {
    return (ast.type === 'Program' ? ast.body.filter(f => f.type === type) : [])
  }

  /**
   * pull out the name of the variable/funciton to be exported
   */
  makeExports(ast, type, accessor) {
    // possible discovered values of type
    // ignoring
    // EmptyExpression
    // ExpressionStatement

    // handling
    // FunctionDeclaration
    // VariableDeclaration
    // ClassDeclaration

    return this.getBody(ast, type).map(c => accessor(c))
  }


  /**
   * compile the files, and get the names of the exports for each file
   * @param {File[]} files the files to be parsed
   * @return {object[]} 
   */
  getExports(params, files) {
    return files.map(this.getAst)
      .map(({ file, ast }) => ({
        file,
        ast,
        functions: this.makeExports(ast, 'FunctionDeclaration', (f) => f.id.name),
        variables: this.makeExports(ast, 'VariableDeclaration', (f) => f.declarations[0].id.name),
        classes: params.exportClasses ? this.makeExports(ast, 'ClassDeclaration', (f) => f.id.name) : []
      }))
  }
  /**
  * @param {[*]} arguments unspecified number and type of args
  * @return {string} a digest of the arguments to use as a key
  */
  keyDigest(...args) {

    // conver args to an array and digest them
    const t = args.map(function (d) {
      return (Object(d) === d) ? JSON.stringify(d) : d.toString();
    }).join("-")
    const s = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, t, Utilities.Charset.UTF_8)
    return Utilities.base64EncodeWebSafe(s)

  };

  getKey({ versionNumber, scriptId }) {
    return `script_${scriptId}_${(versionNumber || '0')}`
  }

  getSourceKey({ source }) {
    return `src_${this.keyDigest(source)}`
  }

  getLibraries(manifest) {
    return (manifest.dependencies && manifest.dependencies.libraries) || []
  }
  /**
   * add scripts to the pool
   */
  addScriptsToPool(params, namespace, scripts) {
    const scriptKey = this.getKey(params)
    if (!scripts.get(scriptKey)) scripts.set(scriptKey, namespace)
    return scripts.get(scriptKey)
  }

  get indenter() {
    const indent = 2
    const t = this.spaces(2)
    const s = this.spaces(indent)
    const sn = s + '\n'
    const ns = '\n' + s
    return {
      indent,
      t,
      s,
      sn,
      ns
    }
  }

  get getterObject() {
    return '__bmlimporter_gets'
  }

  makeInstanceOfLibrary(namespace) {
    const { ns } = this.indenter
    return `${ns}//---Inlined Instance of library ${namespace.scriptId}(${namespace.userSymbol}` +
      ` version ${namespace.versionNumber})` +
      `${ns}const ${namespace.userSymbol} = ${this.getterObject}.${namespace.namespaceName}`
  }

  makeInstanceOfLibraries(children) {
    const { ns } = this.indenter
    return children.map(namespace => this.makeInstanceOfLibrary(namespace)).join(`${ns}`) + ns
  }
  /**
   * make a namespace string from the exports
   * @param {object} params
   * @param {object[]} params.exports the exports object
   * @param {string} params.scriptId the scriptId
   * @param {string} [params.versionNumber] the versionnumber
   * @param {string} params.title project name
   * @param {string} namespaceName what to call the namespace
   * @returns {string} source for the namespace
   */
  generateSource({ exports, namespaceName, title, scriptId, versionNumber, children }) {
    const { indent, t, sn, ns } = this.indenter

    // all the sources of needed to make this export
    const allSource = [].concat(...exports.map(f => this.commenter(f, indent)))
    const allExports = [].concat(...exports.map(f => [].concat(
      f.functions.length ? [`// functions from ${f.file.name}`] : [], ...f.functions,
      f.variables.length ? [`// variables from ${f.file.name}`] : [], ...f.variables,
      f.classes.length ? [`// classes from ${f.file.name}`] : [], ...f.classes)))

    // hoist namespace name
    const hoist = `function ${namespaceName} () {${ns}`

    // the libraries definition
    const libraries = this.makeInstanceOfLibraries(children)

    // the source prettifies
    const code = `${allSource.join('\n').replace(/\r/g, sn)}${ns}`

    // the exports
    const exportProps = `return {${ns + t + allExports.join("," + ns + t)}${ns}}`

    // the whole thing
    const content = [hoist, libraries, code, exportProps, '\n}'].join('')
    const decorated = `${this.namespaceCommenter({ title, scriptId, versionNumber, content })}`
    return decorated
  }



  makeLibraries(params, scripts = new Map(), depth = 0, namespaces = []) {

    // get the tree
    this.getProjectTree(params, scripts, depth, namespaces)

    // assemble the code
    const inlined = namespaces.slice().reverse().map(namespace => {
      const code = namespace.depth ? this.generateSource(namespace) : ''
      return {
        code,
        namespace
      }
    })
    return {
      inlined,
      scripts
    }
  }
  /**
   * get version number to use
   * @param {object} params
   * @param {string} versionTreatment
   * @return {number} version number
   */
  getVersionNumber(params, versionTreatment, { title }) {

    const { scriptId, versionNumber } = params
    const { name, description } = versionTreatment
    const vl = `(${name}:${description})`
    // because it became a string in the manifest
    const vn = versionNumber ? parseInt(versionNumber, 10) : 0

    // don't mess with the given version number
    if (name === 'respect') return vn

    // if development version required
    if (name === 'head') {
      console.log(`Modified reference to ${scriptId}:(${title}) from version ${vn} to head (${vl})`)
      return 0
    }

    // if its an upgrade we need to get the latest deployed version
    if (name !== 'upgrade' && name !== 'sloppyupgrade') {
      throw new Error(`Invalid version treatment ${name}`)
    }

    // get all the deployed versions and sort them in reverse, and take the first one
    const result = this.sapi.listProjectDeployments(params).throw()
    const [deployment] = result.data.deployments
      .filter(f => f.deploymentConfig && f.deploymentConfig.versionNumber)
      .sort((a, b) => b.deploymentConfig.versionNumber - a.deploymentConfig.versionNumber)

    if (deployment) {
      const v = deployment.deploymentConfig.versionNumber

      if (v === vn) {
        console.log(
          `Upgrade not required for reference to library ${scriptId}:(${title}) already specifies version ${v} -  (${vl})`)
      } else {
        // maybe old and new IDE was mixed
        if (vn > v) {
          if (name === 'sloppyupgrade') {
            console.log(`Sloppy upgrade of reference to library ${scriptId}:(${title}) from version ${versionNumber} to head  (${vl})`)
            return 0
          } else {
            throw new Error(`Library ${scriptId}:(${title}) is referenced as later version (${vn}) than the latest deployed (${v}) - maybe ${vn} was deployed with old IDE ? - try a sloppyupgrade`)
          }

        } else {
          console.log(
            `Upgrading reference to library ${scriptId}:(${title}) from version ${versionNumber} to ${v} -  (${vl})`)
        }
      }
      return v
    }

    if (name !== 'sloppyupgrade') {
      throw new Error(`Couldn't find any deployments for ${scriptId}:(${title}) ${vl} - maybe they were deployed with old IDE?`)
    }
    // so return the head if its a sloppyuprade
    console.log(`Sloppy upgrade of reference to library ${scriptId}:(${title}) from version ${versionNumber} to head  (${description})`)
    return 0
  }
  nameSweeper(name) {
    return name.replace(/^[^a-zA-Z_$]|[^\w$]/g, "_")
  }
  /**
   * get project tree
   */
  getProjectTree(params, scripts, depth, namespaces) {

    // inherit sapi params
    const { noCache, userSymbol = '', scriptId, projectContent, exportClasses } = params
    const versionTreatment = this.validateTreatments(params)

    // throw an error on failure as all should be present
    const project = this.sapi.getProject(params).throw()
    console.log(`Working (${depth}) on ${depth ? 'library' : 'project'} ${project.data.title} - ${project.data.scriptId}`)
    if (depth > this.MAX_DEPTH) throw new Error(`Maximum depth ${this.MAX_DEPTH} exceeded - likely circular reference(s) in dependent libraries`)
    // if its a library we're getting, then we need to know which version we're getting
    // if version isnt specified, we'll be getting head anyway
    // if its top level, we dont need to bother
    const versionNumber = depth ? this.getVersionNumber(params, versionTreatment, project.data) : params.versionNumber
    const patchedParams = { ...params, versionNumber }
    // normally we go and fetch the project content, but if the top level content is passed over, we can use that instead (revert/convert would do this)
    const content = projectContent && !depth ? { data: projectContent } : this.sapi.getProjectContent(patchedParams).throw()
    const { data } = content
    const isChild = depth > 0


    // the depth is used to decide whether a namespace is required
    // depth 0 is actually the global space
    const manifest = JSON.parse(this.getManifest(data).source)
    const serverFiles = this.getServers(data)

    const namespace = {
      depth,
      files: data.files.filter(f => !this.sapi.isManifest(f) && f.name !== this.getImportManifestName(patchedParams)),
      namespaceName: this.nameSweeper((isChild ? `__${project.data.title}_${project.data.scriptId.slice(-4)}_v${versionNumber}` : 'global')),
      manifest,
      scriptId,
      versionTreatment: versionTreatment.name,
      versionNumber,
      userSymbol,
      exports: isChild ? this.getExports(params, serverFiles) : null,
      title: project.data.title
    }

    // add the script to the script pool if required
    if (isChild) {
      namespace.script = this.addScriptsToPool(patchedParams, namespace, scripts)
    }
    namespaces.push(namespace)


    // now recurse through all dependent libraries
    namespace.children = this.getLibraries(manifest).map(({ libraryId, version, userSymbol }) => {
      if (libraryId === scriptId) throw new Error(`circular reference to ${scriptId} in ${project.data.title}`)
      const n = this.getProjectTree({
        versionNumber: version,
        scriptId: libraryId,
        userSymbol,
        noCache,
        versionTreatment: versionTreatment.name,
        exportClasses
      }, scripts, depth + 1, namespaces)
      return n
    })
    return namespace

  }

}
// class instance export
var newLimporter = (sapi) => new Limporter(sapi)

// export importscript library
var bmImporter = bmImportScript
