# Collection of my google app scripts

## Usage

1. install clasp
```shell
npm install -g @google/clasp
clasp login
```

2. create project or pull existing project

```shell
# create project
clasp create --title <title> --rootDir <directory>

# pull project
clasp clone <gas script id>
```

3. push source code
```shell
clasp push
```

## Use Typescript

```shell
cp tsconfig.base.json <path>/<to>/<project>/tsconfig.json
mv <project file>.js <project file>.ts
```
