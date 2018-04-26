## spfx-display-group-members

Demo:
![display members](/images/demo.gif)

Remember that the groups must be able to be read (security settings of the group) in order for the web part to display any information.

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.
