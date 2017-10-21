const xlsx = require('xlsx');
const {Client} = require('pg');
/**********************************************CONFIG*********************************************/
const dbConfig = {
    user: 'postgres',
    host: 'localhost',
    database: 'vika',
    password: '',
    port: 5432,
};
const filePath = `${__dirname}/Export.xlsx`;

/**********************************************PARSER*********************************************/
class ExcelParser {
    parseFile(filename) {
        const workbook = xlsx.readFile(filename);
        const sheetName = workbook.SheetNames[0];
        this.sheet = workbook.Sheets[sheetName];

        this.createDataArray();
        this.parseData();
    }

    /**
     * Creates an array of Subjects titles
     */
    createDataArray() {
        const objects = [];
        const keys = Object.keys(this.sheet);
        const lastKey = keys[keys.length - 3];
        const lastKeyNumber = +lastKey.match(/\d+/)[0];
        let lastObject = null;

        for (let i = 4; i < lastKeyNumber; i++) {
            if (this.sheet[`A${i}`] === undefined) {
                if (this.sheet[`B${i}`] !== undefined) {
                    const description = this.sheet[`B${i}`].v;
                    lastObject.descriptions.push(
                        {
                            title: description,
                            idea: lastObject.ideaStandard
                        }
                    );
                }
            } else {
                const object = {
                    benchmark: this.sheet[`A${i}`] ? this.sheet[`A${i}`].v : null,
                    descriptions: this.sheet[`B${i}`] ? [
                        {
                            title: this.sheet[`B${i}`].v,
                            idea: this.sheet[`D${i}`].v
                        }
                    ] : [],
                    ideaStandard: this.sheet[`D${i}`] ? this.sheet[`D${i}`].v : null,
                    subject: this.sheet[`E${i}`] ? this.sheet[`E${i}`].v : null,
                    grade: this.sheet[`F${i}`] ? this.sheet[`F${i}`].v : null,
                    bodyOfKnowledgeStrand: this.sheet[`G${i}`] ? this.sheet[`G${i}`].v : null
                };
                lastObject = object;
                objects.push(object);
            }
        }

        this.data = objects;
    }

    parseData() {
        this.subjects = [];

        for (let row of this.data) {
            const title = `${row.subject}_${row.grade}`;
            let subject = this.findObject(this.subjects, 'title', title);
            if (subject === null) {
                subject = {
                    title: title,
                    titleName: row.subject,
                    grade: row.grade,
                    bodies: []
                };
                this.subjects.push(subject);
            }

            let body = this.findObject(subject.bodies, 'title', row.bodyOfKnowledgeStrand);
            if (body === null) {
                body = {
                    title: row.bodyOfKnowledgeStrand,
                    ideas: [],
                    benchmark: row.benchmark
                };
                subject.bodies.push(body);
            }

            let idea = this.findObject(body.ideas, 'title', row.ideaStandard);
            if (idea === null) {
                idea = {
                    title: row.ideaStandard,
                    descriptions: [],
                    benchmark: row.benchmark
                };
                body.ideas.push(idea);
            }
            for (let description of row.descriptions) {
                if (description.idea === idea.title) {
                    idea.descriptions.push({
                        title: description.title,
                        benchmark: row.benchmark
                    });
                }
            }
        }
    }

    findObject(array, key, value) {
        for (let row of array) {
            if (key !== null && row[key] === value) {
                return row;
            }

            if (key === null && row === value) {
                return row;
            }
        }

        return null;
    }
}

/**********************************************FORMATTER******************************************/
class DataFormatter {
    constructor(subjects) {
        this.subjects = subjects;
    }

    formatData() {
        this.formatSubjects();
        this.formatBodies();
        this.formatIdeas();
        this.formatDescription();
    }

    formatBodies() {
        for (let subject of this.subjects) {
            for (let body of subject.bodies) {
                body.code = body.benchmark.split('.')[2];
            }

        }
    }

    formatIdeas() {
        for (let subject of this.subjects) {
            for (let body of subject.bodies) {
                for (let idea of body.ideas) {
                    idea.code = idea.benchmark.split('.')[3];
                }
            }

        }
    }

    formatDescription() {
        for (let subject of this.subjects) {
            for (let body of subject.bodies) {
                for (let idea of body.ideas) {
                    for (let description of idea.descriptions) {
                        description.code = description.benchmark.split('.')[4];

                        if (description.title[1] === '.' && isNaN(description.title[0])) {
                            description.code += "." + description.title[0];
                            description.title = description.title.substr(2);

                        }
                    }
                }
            }

        }
    }

    formatSubjects() {
        for (let subject of this.subjects) {
            if (!isNaN(subject.grade)) {
                subject.gradeText = this.formatNumber(subject.grade);
            } else {
                subject.gradeText = this.formatString(subject.grade);
            }

            const gradeText = subject.gradeText + '';

            subject.grades = this.createGradeFromGradeText(subject.gradeText);
            subject.gradeText = this.reformatGradeText(subject.gradeText);

            if (~gradeText.indexOf('-')) {
                subject.titleName += " - Grades " + subject.gradeText;
            } else {
                subject.titleName += " - Grade " + subject.gradeText;
            }
        }
    }

    formatNumber(numbers) {
        if (numbers > 1000) {
            numbers = numbers + "";

            return numbers.slice(0, 2) + "-" + numbers.slice(2);
        }

        if (numbers > 12) {
            numbers = numbers + "";

            return numbers.charAt(0) + "-" + numbers.slice(1);
        }

        return numbers;
    }

    formatString(strings) {
        if (strings === "K") {
            return "KG";
        }

        if (strings === "P") {
            return "PK";
        }

        if (strings.length > 1 && !isNaN(strings.charAt(1))) {
            let numbers = this.formatNumber(strings.slice(1));

            return this.formatString(strings.charAt(0)) + "," + numbers;
        }
        if (strings.length > 1 && isNaN(strings.charAt(1))) {
            let numbers = this.formatNumber(strings.slice(2));
            return strings.slice(0, 2) + "," + numbers;
        }
    }

    createGradeFromGradeText(gradeText) {
        gradeText = gradeText + '';
        gradeText = gradeText.split(',');
        let arrayLimitedGrades = [];

        for (let value of gradeText) {
            if (~value.indexOf('-')) {
                let limit = value.split('-');
                for (let i = +limit[0]; i <= +limit[1]; i++) {
                    arrayLimitedGrades.push(i);
                }
            }
            if (!isNaN(value) && +value > 1) {
                console.log(value, '--------------------------------------------');
                for (let i = 0; i < value; i++) {
                    arrayLimitedGrades.push(i + 1);
                }
                console.log(arrayLimitedGrades, '_______________________________________-');

            }
        }
        return arrayLimitedGrades;
    }

    reformatGradeText(value) {
        for (let i = 0; i < value.length; i++) {
            if (value.charAt(i) === ',' && isNaN(value.charAt(i - 1)) && !isNaN(value.charAt(i + 1))) {
                return value.substr(0, i) + "-" + value.substr(i + 1);
            }
        }

        return value;
    }

}

/**********************************************DB INSERTER****************************************/
class DataInserter {
    constructor(subjects) {
        this.connection = new Client(dbConfig);
        this.connection.connect();
        this.subjects = subjects;
    }

    insertAll() {
        this.connection.query("INSERT INTO organization(title) VALUES  ( 'Florida State Standards') ");
        this.connection.query("SELECT id FROM organization WHERE title = 'Florida State Standards' ")
            .then((res) => {
                const idOfOrganisation = res.rows[0].id;

                this.insertSubjects(idOfOrganisation);
            }).catch((e) => {
            console.log(e)
        });
    }

    insertSubjects(idOfOrganisation) {
        const insertSubjectQuery = "INSERT INTO standard(title, gradetext, grades, organizationid) VALUES ($1, $2, $3, $4)";

        for (const subject of parser.subjects) {
            const values = [subject.titleName, subject.gradeText, JSON.stringify(subject.grades), idOfOrganisation];
            this.connection.query(insertSubjectQuery, values);

            this.connection.query(`SELECT id FROM standard WHERE title = '${subject.titleName}' `)
                .then((res) => {
                    subject.id = res.rows[0].id;

                    this.insertBodies(subject);
                }).catch(() => {
            });
        }
    }

    insertBodies(subject) {
        const insertSubjectQuery = "INSERT INTO keyIdea(title, code, standardid) VALUES ($1, $2, $3)";

        for (const body of subject.bodies) {
            const values = [body.title, body.code, subject.id];
            this.connection.query(insertSubjectQuery, values);

            this.connection.query(`SELECT id FROM keyIdea WHERE title = '${body.title}' AND standardid = '${subject.id}'`)
                .then((res) => {
                    body.id = res.rows[0].id;
                    this.insertDomain(body);
                }).catch(() => {
            });
        }
    }

    insertDomain(body) {
        const insertSubjectQuery = "INSERT INTO domain(title, code, keyideaid) VALUES ($1, $2, $3)";

        const values = [body.title, body.code, body.id];
        this.connection.query(insertSubjectQuery, values);

        this.connection.query(`SELECT id FROM domain WHERE keyideaid = '${body.id}'`)
            .then((res) => {
                body.id = res.rows[0].id;
                this.insertIdeas(body);
            }).catch(() => {
        });
    }

    insertIdeas(domain) {
        const insertSubjectQuery = "INSERT INTO statement(title, code, domainid) VALUES ($1, $2, $3)";

        for (const idea of domain.ideas) {
            const values = [idea.title, idea.code, domain.id];
            this.connection.query(insertSubjectQuery, values);

            this.connection.query(`SELECT id FROM statement WHERE title = '${idea.title}' AND domainid = '${domain.id}'`)
                .then((res) => {
                    idea.id = res.rows[0].id;
                    this.insertDescriptions(idea);
                }).catch(() => {
            });
        }
    }

    insertDescriptions(idea) {
        const insertSubjectQuery = "INSERT INTO expectation(title, code, statementid) VALUES ($1, $2, $3)";

        for (const description of idea.descriptions) {
            let desc = description.title;
            desc = desc.replace('²', '^2');
            desc = desc.replace('³', '^3');
            desc = desc.replace('★', '*');
            desc = desc.replace(/√\d+/, function (match) {
                return 'sqrt(' + match.substr(1) + ')';
            });

            const values = [desc, description.code, idea.id];
            this.connection.query(insertSubjectQuery, values).then(() => 2);
        }
    }
}

/**********************************************EXECUTION******************************************/
// Execution
const parser = new ExcelParser();
parser.parseFile(filePath);
const formatter = new DataFormatter(parser.subjects);
formatter.formatData();
const inserter = new DataInserter(formatter.subjects);
inserter.insertAll();