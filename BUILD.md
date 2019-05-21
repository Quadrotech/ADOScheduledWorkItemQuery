# Prerequisites

- node.js/npm
- Typescript
- Azure DevOps Extension Publish Tool (npm i -g tfx-cli)

# Build

```
cd ado-scheduledworkitemquery-taskV1
tsc  (TypeScript compile)

cd ..
tfx extension create --manifest-globs vss-extension.json vss-extension-test.json
```

