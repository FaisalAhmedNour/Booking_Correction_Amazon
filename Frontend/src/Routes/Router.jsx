import { createMemoryRouter } from "react-router-dom";
import Main from "../Layouts/Main/Main";
import DownloadPO from "../pages/DownloadPO/DownloadPO";
import BookingBuilder from "../pages/BookingBuilder/BookingBuilder";

export const router = createMemoryRouter([
    {
        path: "/",
        element: <Main />,
        children: [
            {
                path: "/",
                element: <DownloadPO />
            },
            {
                path: "/builder",
                element: <BookingBuilder />
            }
        ]
    }
]);