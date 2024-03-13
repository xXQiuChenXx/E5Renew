import  { Request, Response  } from 'express';

const middleware = async (req:Request, res: Response) => {
    console.log("middleware")
}

export default middleware;